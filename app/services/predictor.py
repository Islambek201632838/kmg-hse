import pandas as pd
import numpy as np
from functools import lru_cache
from app.services.data_loader import load_incidents, load_korgau
from app.services.korgau_analyzer import get_org_trends

# ─────────────────────────────────────────────
# Справочник мер контроля → категория Коргау + ожидаемое снижение нарушений
# ─────────────────────────────────────────────
CONTROL_MEASURES: dict[str, dict] = {
    "LOTO (блокировка/тегирование)": {"category": "LOTO", "reduction_pct": 70},
    "Обучение по СИЗ": {"category": "СИЗ", "reduction_pct": 60},
    "Инструктаж по работе на высоте": {"category": "Работа на высоте", "reduction_pct": 65},
    "Пожарный аудит": {"category": "Пожарная безопасность", "reduction_pct": 55},
    "Безопасность транспорта": {"category": "Транспорт", "reduction_pct": 55},
    "Электробезопасность": {"category": "Электробезопасность", "reduction_pct": 60},
    "Химическая безопасность": {"category": "Химическая безопасность", "reduction_pct": 50},
    "Порядок на рабочем месте (5S)": {"category": "Организация рабочего места", "reduction_pct": 50},
    "Система разрешений (ПТР)": {"category": "Процедуры и разрешения", "reduction_pct": 65},
}


# ─────────────────────────────────────────────
# Prophet forecast
# ─────────────────────────────────────────────

@lru_cache(maxsize=1)
def get_forecast(periods: int = 12) -> dict:
    """Prophet forecast. При недоступности cmdstan — линейный fallback."""
    df = load_incidents()

    monthly = (
        df.groupby(df["_date"].dt.to_period("M"))
        .size()
        .reset_index(name="y")
    )
    monthly["ds"] = monthly["_date"].dt.to_timestamp()
    monthly = monthly[["ds", "y"]].sort_values("ds")

    if len(monthly) < 3:
        return {"error": "Недостаточно данных для прогноза"}

    try:
        from prophet import Prophet
        model = Prophet(
            yearly_seasonality=True,
            weekly_seasonality=False,
            daily_seasonality=False,
            interval_width=0.95,
        )
        model.fit(monthly)
        future = model.make_future_dataframe(periods=periods, freq="MS")
        forecast = model.predict(future)

        history = forecast[forecast["ds"].isin(monthly["ds"])][
            ["ds", "yhat", "yhat_lower", "yhat_upper"]
        ].copy()
        history = history.merge(monthly, on="ds", how="left")

        last_hist_ds = monthly["ds"].max()
        future_fc = forecast[forecast["ds"] > last_hist_ds][
            ["ds", "yhat", "yhat_lower", "yhat_upper"]
        ].copy()
        future_fc["yhat"] = future_fc["yhat"].clip(lower=0).round(1)
        future_fc["yhat_lower"] = future_fc["yhat_lower"].clip(lower=0).round(1)
        future_fc["yhat_upper"] = future_fc["yhat_upper"].clip(lower=0).round(1)
        engine = "prophet"

    except Exception:
        # Линейный тренд как fallback если cmdstan не установлен
        history = monthly.copy()
        history["yhat"] = history["y"].astype(float)
        history["yhat_lower"] = history["yhat"]
        history["yhat_upper"] = history["yhat"]

        # Простая линейная экстраполяция
        x = np.arange(len(monthly))
        y = monthly["y"].values.astype(float)
        slope, intercept = np.polyfit(x, y, 1)

        last_ds = monthly["ds"].max()
        future_dates = pd.date_range(
            start=last_ds + pd.DateOffset(months=1), periods=periods, freq="MS"
        )
        future_y = np.array([
            max(0, slope * (len(monthly) + i) + intercept)
            for i in range(periods)
        ])
        future_fc = pd.DataFrame({
            "ds": future_dates,
            "yhat": future_y.round(1),
            "yhat_lower": (future_y * 0.7).round(1),
            "yhat_upper": (future_y * 1.3).round(1),
        })
        engine = "linear_fallback"

    # Компоненты тренда (только у Prophet)
    if engine == "prophet":
        trend_component = forecast[["ds", "trend"]].copy()
        trend_component["trend"] = trend_component["trend"].round(2)
        trend_data = trend_component.to_dict("records")
    else:
        trend_data = []

    # Backtesting: MAE / RMSE на исторических данных (fitted vs actual)
    hist_with_y = history.dropna(subset=["y", "yhat"])
    if len(hist_with_y) > 0:
        residuals = hist_with_y["y"] - hist_with_y["yhat"]
        mae = round(float(residuals.abs().mean()), 2)
        rmse = round(float(np.sqrt((residuals ** 2).mean())), 2)
        mape = round(float((residuals.abs() / hist_with_y["y"].replace(0, 1)).mean() * 100), 1)
    else:
        mae = rmse = mape = None

    return {
        "history": history[["ds", "y", "yhat", "yhat_lower", "yhat_upper"]].to_dict("records"),
        "forecast": future_fc[["ds", "yhat", "yhat_lower", "yhat_upper"]].to_dict("records"),
        "trend": trend_data,
        "periods": periods,
        "total_historical_months": len(monthly),
        "engine": engine,
        "metrics": {"mae": mae, "rmse": rmse, "mape_pct": mape},
    }


# ─────────────────────────────────────────────
# Risk Scoring
# ─────────────────────────────────────────────

@lru_cache(maxsize=1)
def get_risk_scores() -> list[dict]:
    """
    risk_score = 0.35 * incident_rate_norm
               + 0.30 * violation_rate_norm
               + 0.20 * trend_slope_norm
               + 0.15 * severity_factor_norm
    Нормализация 0–100 по всем организациям.
    """
    df_inc = load_incidents()
    df_kor = load_korgau()
    org_trends = {r["org"]: r for r in get_org_trends()}

    # ── 1. Incident rate (инциденты на организацию) ──
    inc_counts = df_inc.groupby("_org").size().rename("incident_count")

    # ── 2. Violation rate (нарушения Коргау на организацию) ──
    viol_counts = (
        df_kor[df_kor["_is_violation"]]
        .groupby("_org")
        .size()
        .rename("violation_count")
    )

    # ── 3. Severity factor (вес по тяжести инцидентов) ──
    severity_weights = {
        "Летальный случай (погиб)": 4.0,
        "Относится к тяжелым": 2.0,
        "Не относится к тяжелым": 1.0,
        "Нет данных": 0.5,
    }
    df_inc["_sev_weight"] = df_inc["_severity"].map(severity_weights).fillna(0.5)
    sev_scores = df_inc.groupby("_org")["_sev_weight"].mean().rename("severity_factor")

    # ── 4. Trend slope (% изменение нарушений) ──
    trend_slopes = pd.Series(
        {org: abs(data["trend_pct"]) for org, data in org_trends.items()},
        name="trend_slope",
    )

    # Объединяем все компоненты
    all_orgs = inc_counts.index.union(viol_counts.index)
    scores = pd.DataFrame(index=all_orgs)
    scores = scores.join(inc_counts).join(viol_counts).join(sev_scores).join(trend_slopes)
    scores = scores.fillna(0)

    # Нормализация каждого компонента 0–100
    def normalize(series: pd.Series) -> pd.Series:
        mn, mx = series.min(), series.max()
        if mx == mn:
            return pd.Series(0.0, index=series.index)
        return (series - mn) / (mx - mn) * 100

    scores["incident_rate_norm"] = normalize(scores["incident_count"])
    scores["violation_rate_norm"] = normalize(scores["violation_count"])
    scores["trend_slope_norm"] = normalize(scores["trend_slope"])
    scores["severity_factor_norm"] = normalize(scores["severity_factor"])

    # Итоговый риск-скор
    scores["risk_score"] = (
        0.35 * scores["incident_rate_norm"]
        + 0.30 * scores["violation_rate_norm"]
        + 0.20 * scores["trend_slope_norm"]
        + 0.15 * scores["severity_factor_norm"]
    ).round(1)

    # Уровень риска
    def risk_level(score: float) -> str:
        if score >= 70:
            return "CRITICAL"
        elif score >= 50:
            return "HIGH"
        elif score >= 25:
            return "MEDIUM"
        return "LOW"

    scores["risk_level"] = scores["risk_score"].apply(risk_level)

    result = scores.reset_index().rename(columns={"_org": "org"})
    return (
        result[["org", "risk_score", "risk_level",
                "incident_count", "violation_count",
                "severity_factor", "trend_slope",
                "incident_rate_norm", "violation_rate_norm",
                "trend_slope_norm", "severity_factor_norm"]]
        .sort_values("risk_score", ascending=False)
        .to_dict("records")
    )


# ─────────────────────────────────────────────
# Корреляционная матрица
# ─────────────────────────────────────────────

@lru_cache(maxsize=1)
def get_correlation_matrix() -> dict:
    """Матрица корреляций: категории нарушений Коргау × типы инцидентов."""
    df_kor = load_korgau()
    df_inc = load_incidents()

    # Пивот: (org, month) × категория нарушений
    kor_pivot = (
        df_kor[df_kor["_is_violation"]]
        .assign(month=df_kor["_date"].dt.to_period("M"))
        .groupby(["_org", "month", "_category"])
        .size()
        .unstack(fill_value=0)
    )

    # Пивот: (org, month) × тип инцидента
    inc_pivot = (
        df_inc
        .assign(month=df_inc["_date"].dt.to_period("M"))
        .groupby(["_org", "month", "_type"])
        .size()
        .unstack(fill_value=0)
    )

    # Объединяем по (org, month), суффиксы на случай совпадений имён
    merged = kor_pivot.join(inc_pivot, how="inner", lsuffix="_kor", rsuffix="_inc")
    # Восстанавливаем оригинальные списки колонок с учётом суффиксов
    kor_cols = [c + "_kor" if c in inc_pivot.columns else c for c in kor_pivot.columns]
    inc_cols = [c + "_inc" if c in kor_pivot.columns else c for c in inc_pivot.columns]
    if merged.empty or len(merged) < 3:
        return {"matrix": [], "categories": [], "incident_types": []}

    kor_cols = list(kor_pivot.columns)
    inc_cols = list(inc_pivot.columns)

    # Считаем корреляцию каждой категории нарушений с каждым типом инцидента
    matrix = {}
    for kor_cat in kor_cols:
        matrix[kor_cat] = {}
        for inc_type in inc_cols:
            if kor_cat in merged.columns and inc_type in merged.columns:
                corr = merged[kor_cat].corr(merged[inc_type])
                matrix[kor_cat][inc_type] = round(float(corr) if not np.isnan(corr) else 0.0, 3)

    # Топ-10 категорий и все типы инцидентов для UI
    top_kor_cats = (
        df_kor[df_kor["_is_violation"]]["_category"]
        .value_counts()
        .head(10)
        .index.tolist()
    )
    filtered = {k: v for k, v in matrix.items() if k in top_kor_cats}

    return {
        "matrix": filtered,
        "categories": top_kor_cats,
        "incident_types": inc_cols,
    }


# ─────────────────────────────────────────────
# Scenario Modeling (3.1.2)
# ─────────────────────────────────────────────

def calculate_scenario(org: str, measures: list[str]) -> dict:
    """
    Расчёт изменения вероятности инцидентов при внедрении мер контроля.

    Методология:
    1. Каждая мера снижает нарушения в соответствующей категории Коргау.
    2. Вклад категории взвешивается по её доле в общем числе нарушений org.
    3. Итоговое снижение нарушений умножается на Pearson r=0.415
       (корреляция нарушений → инциденты), что даёт ожидаемое снижение инцидентов.
    4. Экономический эффект: избежанные инциденты × 5 млн ₸ (прямые затраты на НС).
    """
    df_inc = load_incidents()
    df_kor = load_korgau()

    # Pearson-корреляция нарушений с инцидентами (установлена по всей выборке)
    PEARSON_R = 0.415

    # Базовые инциденты по организации
    org_incidents = df_inc[df_inc["_org"] == org]
    baseline = len(org_incidents)

    # Нарушения по категориям для организации
    org_viol = df_kor[(df_kor["_org"] == org) & (df_kor["_is_violation"])]
    viol_by_cat = org_viol["_category"].value_counts()
    total_viol = viol_by_cat.sum() if not viol_by_cat.empty else 0

    measures_applied = []
    total_violation_reduction_weighted = 0.0

    for measure in measures:
        if measure not in CONTROL_MEASURES:
            continue
        m = CONTROL_MEASURES[measure]
        cat = m["category"]
        red_pct = m["reduction_pct"]

        cat_viol = int(viol_by_cat.get(cat, 0))
        # Вес категории в общем портфеле нарушений org (если данных нет — берём 5% базово)
        weight = (cat_viol / total_viol) if total_viol > 0 else 0.05

        # Взвешенный вклад в снижение нарушений
        weighted_contrib = weight * red_pct
        total_violation_reduction_weighted += weighted_contrib

        # Вклад в снижение инцидентов = violation_reduction × Pearson r
        incident_impact_pct = round(weighted_contrib * PEARSON_R, 1)

        measures_applied.append({
            "measure": measure,
            "category": cat,
            "current_violations": cat_viol,
            "violation_reduction_pct": red_pct,
            "incident_impact_pct": incident_impact_pct,
        })

    # Итоговое снижение инцидентов (не более 70%)
    violation_reduction_pct = min(total_violation_reduction_weighted, 100.0)
    incident_reduction_pct = round(min(violation_reduction_pct * PEARSON_R, 70.0), 1)

    predicted = round(baseline * (1 - incident_reduction_pct / 100), 1)
    avoided = round(baseline - predicted, 1)

    # Экономический эффект: 60% инцидентов — НС, прямые затраты 5 млн ₸/НС
    economic_saving = int(avoided * 0.6 * 5_000_000)

    return {
        "org": org,
        "baseline_incidents": baseline,
        "predicted_incidents": predicted,
        "avoided_incidents": avoided,
        "incident_reduction_pct": incident_reduction_pct,
        "violation_reduction_pct": round(violation_reduction_pct, 1),
        "economic_saving_tenge": economic_saving,
        "measures_applied": measures_applied,
        "available_measures": list(CONTROL_MEASURES.keys()),
    }
