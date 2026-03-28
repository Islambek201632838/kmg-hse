import pandas as pd
import numpy as np
from scipy import stats
from datetime import datetime, timedelta
from app.services.data_loader import load_korgau, load_incidents

# Пороги алертов (нарушений за последние 30 дней по организации)
ALERT_THRESHOLDS = {
    "CRITICAL": 50,
    "HIGH":     30,
    "MEDIUM":   15,
    "LOW":       5,
}


def get_violation_rates() -> list[dict]:
    """Violation rate по (организация, категория, месяц)."""
    df = load_korgau()
    violations = df[df["_is_violation"]]

    grouped = (
        violations
        .groupby(["_org", "_category", violations["_date"].dt.to_period("M")])
        .size()
        .reset_index(name="count")
    )
    grouped["month"] = grouped["_date"].dt.to_timestamp()
    grouped.drop(columns=["_date"], inplace=True)
    return grouped.to_dict("records")


def get_org_trends() -> list[dict]:
    """Тренд нарушений по каждой организации: рост/снижение %."""
    df = load_korgau()
    violations = df[df["_is_violation"]]

    result = []
    for org, group in violations.groupby("_org"):
        monthly = group.groupby(group["_date"].dt.to_period("M")).size().sort_index()

        total = len(group)
        trend_pct = 0.0
        direction = "стабильно"

        if len(monthly) >= 4:
            recent = monthly.iloc[-2:].mean()
            old = monthly.iloc[-4:-2].mean()
            if old > 0:
                trend_pct = round((recent - old) / old * 100, 1)
                direction = "рост" if trend_pct > 5 else ("снижение" if trend_pct < -5 else "стабильно")

        result.append({
            "org": org,
            "total_violations": total,
            "trend_pct": trend_pct,
            "direction": direction,
            "monthly": {str(k): int(v) for k, v in monthly.tail(12).items()},
        })

    return sorted(result, key=lambda x: x["total_violations"], reverse=True)


def get_alerts(reference_date: datetime | None = None) -> list[dict]:
    """4-уровневые алерты по количеству нарушений за последние 30 дней."""
    df = load_korgau()
    violations = df[df["_is_violation"]]

    if reference_date is None:
        reference_date = violations["_date"].max()

    window_start = reference_date - timedelta(days=30)
    recent = violations[violations["_date"] >= window_start]

    counts = recent.groupby("_org").size().reset_index(name="count_30d")

    alerts = []
    for _, row in counts.iterrows():
        n = row["count_30d"]
        if n > ALERT_THRESHOLDS["CRITICAL"]:
            level = "CRITICAL"
            color = "red"
        elif n > ALERT_THRESHOLDS["HIGH"]:
            level = "HIGH"
            color = "orange"
        elif n > ALERT_THRESHOLDS["MEDIUM"]:
            level = "MEDIUM"
            color = "yellow"
        elif n > ALERT_THRESHOLDS["LOW"]:
            level = "LOW"
            color = "blue"
        else:
            continue  # Ниже порога — не показываем

        # Топ категории нарушений для этой организации за период
        org_violations = recent[recent["_org"] == row["_org"]]
        top_categories = org_violations["_category"].value_counts().head(3).to_dict()

        # F-11: один тип нарушения >3 раз за 30 дней — информация (не повышаем уровень)
        category_counts = org_violations["_category"].value_counts()
        repeated_reason = None
        if (category_counts > 3).any():
            top_cat = category_counts.idxmax()
            repeated_reason = f"{top_cat} — {int(category_counts.max())} раз за 30 дней"

        # F-14 (MEDIUM): YoY рост нарушений >15% — повышаем только если уровень LOW
        yoy_reason = None
        org_all = violations[violations["_org"] == row["_org"]]
        if len(org_all) > 0:
            year_ago = reference_date - timedelta(days=365)
            cur_count = len(org_all[org_all["_date"] >= year_ago])
            prev_count = len(org_all[(org_all["_date"] >= year_ago - timedelta(days=365)) & (org_all["_date"] < year_ago)])
            if prev_count > 0:
                yoy_pct = round((cur_count - prev_count) / prev_count * 100, 1)
                if yoy_pct > 15:
                    yoy_reason = f"YoY рост: {yoy_pct}% ({prev_count} → {cur_count})"

        alerts.append({
            "org": row["_org"],
            "level": level,
            "color": color,
            "count_30d": int(n),
            "top_categories": top_categories,
            "period": f"{window_start.date()} — {reference_date.date()}",
            "repeated_violation_reason": repeated_reason,
            "yoy_reason": yoy_reason,
        })

    return sorted(alerts, key=lambda x: x["count_30d"], reverse=True)


def get_org_rankings() -> list[dict]:
    """Ранжирование организаций по уровню соответствия/несоответствия."""
    df = load_korgau()

    result = []
    for org, group in df.groupby("_org"):
        total = len(group)
        violations = group[group["_is_violation"]]
        good = group[~group["_is_violation"]]

        violation_rate = round(len(violations) / max(total, 1) * 100, 1)
        compliance_rate = round(100 - violation_rate, 1)

        top_violation_categories = (
            violations["_category"].value_counts().head(3).to_dict()
        )
        work_stops = (violations["_work_stopped"] == "Да").sum()

        result.append({
            "org": org,
            "total_observations": total,
            "violation_count": len(violations),
            "good_practice_count": len(good),
            "violation_rate_pct": violation_rate,
            "compliance_rate_pct": compliance_rate,
            "work_stops": int(work_stops),
            "top_violation_categories": top_violation_categories,
        })

    return sorted(result, key=lambda x: x["violation_rate_pct"], reverse=True)


def get_good_vs_bad() -> dict:
    """Хорошие vs плохие практики по категориям."""
    df = load_korgau()

    good = (
        df[~df["_is_violation"]]
        .groupby("_category")
        .size()
        .sort_values(ascending=False)
        .head(10)
        .to_dict()
    )
    bad = (
        df[df["_is_violation"]]
        .groupby("_category")
        .size()
        .sort_values(ascending=False)
        .head(10)
        .to_dict()
    )
    return {"good": good, "bad": bad}


def get_correlation_with_incidents() -> dict:
    """Pearson/Spearman корреляция между нарушениями Коргау и инцидентами по организациям."""
    korgau = load_korgau()
    incidents = load_incidents()

    # Агрегируем по (org, month)
    kor_monthly = (
        korgau[korgau["_is_violation"]]
        .groupby(["_org", korgau["_date"].dt.to_period("M")])
        .size()
        .reset_index(name="violations")
    )
    inc_monthly = (
        incidents
        .groupby(["_org", incidents["_date"].dt.to_period("M")])
        .size()
        .reset_index(name="incidents")
    )

    merged = pd.merge(
        kor_monthly, inc_monthly,
        on=["_org", "_date"],
        how="inner"
    )

    if len(merged) < 5:
        return {"pearson": None, "spearman": None, "n_points": len(merged)}

    pearson_r, pearson_p = stats.pearsonr(merged["violations"], merged["incidents"])
    spearman_r, spearman_p = stats.spearmanr(merged["violations"], merged["incidents"])

    # Корреляция по организациям (cross-sectional)
    org_agg = merged.groupby("_org").agg(
        violations=("violations", "sum"),
        incidents=("incidents", "sum")
    ).reset_index()

    org_pearson = None
    if len(org_agg) >= 3:
        org_pearson_r, _ = stats.pearsonr(org_agg["violations"], org_agg["incidents"])
        org_pearson = round(float(org_pearson_r), 3)

    # Scatter data — по месяцам для организаций, которые есть в обоих датасетах
    # (outer join давал мусор: орги без пересечения)
    org_scatter = merged[["_org", "violations", "incidents"]].copy()
    org_scatter["_org"] = org_scatter["_org"].astype(str)
    # Добавляем номер месяца для подписи
    org_scatter["label"] = org_scatter["_org"].str[:20] + " (мес " + (org_scatter.index + 1).astype(str) + ")"

    return {
        "pearson_r": round(float(pearson_r), 3),
        "pearson_p": round(float(pearson_p), 4),
        "spearman_r": round(float(spearman_r), 3),
        "spearman_p": round(float(spearman_p), 4),
        "n_points": len(merged),
        "org_level_pearson": org_pearson,
        "scatter_data": org_scatter[["_org", "label", "violations", "incidents"]].to_dict("records"),
    }
