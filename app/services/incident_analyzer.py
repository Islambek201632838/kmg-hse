import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
from app.services.data_loader import load_incidents


def get_summary() -> dict:
    df = load_incidents()

    by_type = df["_type"].value_counts().to_dict()
    by_org = df["_org"].value_counts().head(10).to_dict()
    by_location = df["_location"].value_counts().head(10).to_dict()
    by_severity = df["_severity"].value_counts().to_dict()

    # Ежемесячный тренд + скользящее среднее 3м
    monthly = (
        df.groupby(df["_date"].dt.to_period("M"))
        .size()
        .reset_index(name="count")
    )
    monthly["ds"] = monthly["_date"].dt.to_timestamp()
    monthly["rolling_3m"] = monthly["count"].rolling(3, min_periods=1).mean().round(1)

    # Month-over-month изменение
    if len(monthly) >= 2:
        last = monthly["count"].iloc[-1]
        prev = monthly["count"].iloc[-2]
        mom_change = round((last - prev) / max(prev, 1) * 100, 1)
    else:
        mom_change = 0.0

    # По дням недели и часам (паттерны времени)
    by_weekday = (
        df["_date"].dt.day_name().value_counts().to_dict()
    )

    return {
        "total": len(df),
        "by_type": by_type,
        "by_org": by_org,
        "by_location": by_location,
        "by_severity": by_severity,
        "by_weekday": by_weekday,
        "monthly_trend": monthly[["ds", "count", "rolling_3m"]].tail(24).to_dict("records"),
        "mom_change_pct": mom_change,
    }


def get_cause_clusters(n_clusters: int = 6) -> list[dict]:
    df = load_incidents()
    causes = df["_cause"].dropna().astype(str)
    causes = causes[causes.str.strip().str.len() > 3]

    if len(causes) < n_clusters:
        return []

    vectorizer = TfidfVectorizer(max_features=500, min_df=2, ngram_range=(1, 2))
    X = vectorizer.fit_transform(causes)

    km = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
    labels = km.fit_predict(X)

    feature_names = vectorizer.get_feature_names_out()
    clusters = []
    for i in range(n_clusters):
        top_idx = km.cluster_centers_[i].argsort()[-5:][::-1]
        top_terms = [feature_names[j] for j in top_idx]
        cluster_causes = causes[labels == i]
        count = len(cluster_causes)
        # Сэмпл реальных причин для UI
        examples = cluster_causes.head(3).tolist()
        clusters.append({
            "cluster_id": i,
            "top_terms": top_terms,
            "label": " / ".join(top_terms[:3]),
            "count": count,
            "examples": examples,
        })

    # Silhouette Score — качество кластеризации (-1..1, чем выше тем лучше)
    from sklearn.metrics import silhouette_score, davies_bouldin_score
    sil = round(float(silhouette_score(X, labels)), 3) if len(set(labels)) > 1 else 0.0
    db = round(float(davies_bouldin_score(X.toarray(), labels)), 3) if len(set(labels)) > 1 else 0.0

    result = sorted(clusters, key=lambda x: x["count"], reverse=True)
    # Добавляем метрики в первый элемент (или отдельно)
    for c in result:
        c["silhouette_score"] = sil
        c["davies_bouldin_score"] = db
    return result


def get_patterns() -> dict:
    df = load_incidents()

    top_orgs = df["_org"].value_counts().head(5).index.tolist()
    top_types = df["_type"].value_counts().head(5).index.tolist()

    # Топ причин — по частоте встречаемости строк
    top_causes = (
        df["_cause"]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .value_counts()
        .head(5)
        .index.tolist()
    )

    # Тренд: сравниваем последние 3 мес с предыдущими 3 мес
    monthly = df.groupby(df["_date"].dt.to_period("M")).size()
    if len(monthly) >= 6:
        recent = monthly.iloc[-3:].mean()
        old = monthly.iloc[-6:-3].mean()
        trend_pct = round((recent - old) / max(old, 1) * 100, 1)
    else:
        trend_pct = 0.0

    return {
        "top_orgs": top_orgs,
        "top_types": top_types,
        "top_causes": top_causes,
        "trend": f"{trend_pct:+.1f}%",
        "trend_pct": trend_pct,
    }


def get_org_breakdown() -> list[dict]:
    """Детализация по организациям: кол-во, типы, тренд."""
    df = load_incidents()
    result = []
    for org, group in df.groupby("_org"):
        monthly = group.groupby(group["_date"].dt.to_period("M")).size()
        trend = 0.0
        if len(monthly) >= 4:
            trend = round(
                (monthly.iloc[-2:].mean() - monthly.iloc[-4:-2].mean())
                / max(monthly.iloc[-4:-2].mean(), 1) * 100, 1
            )
        result.append({
            "org": org,
            "total": len(group),
            "by_type": group["_type"].value_counts().to_dict(),
            "trend_pct": trend,
        })
    return sorted(result, key=lambda x: x["total"], reverse=True)
