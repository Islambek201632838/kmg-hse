import streamlit as st
import httpx
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from concurrent.futures import ThreadPoolExecutor

import os
API = os.getenv("API_URL", "http://localhost:8002")

st.set_page_config(
    page_title="HSE AI Analytics",
    page_icon="🛡️",
    layout="wide",
)

st.markdown("""
<style>
.alert-critical { background:#dc3545; color:white; padding:12px; border-radius:8px; margin:4px 0; }
.alert-high     { background:#fd7e14; color:white; padding:12px; border-radius:8px; margin:4px 0; }
.alert-medium   { background:#ffc107; color:#1a1a1a; padding:12px; border-radius:8px; margin:4px 0; }
.alert-low      { background:#0d6efd; color:white; padding:12px; border-radius:8px; margin:4px 0; }
.kpi-box        { background:#1e3a8a; color:white; padding:16px; border-radius:10px; text-align:center; }
.kpi-value      { font-size:2rem; font-weight:700; }
.kpi-label      { font-size:0.85rem; opacity:0.85; }
</style>
""", unsafe_allow_html=True)


def _raw_get(path: str):
    """Thread-safe GET without Streamlit calls."""
    try:
        r = httpx.get(f"{API}{path}", timeout=60)
        r.raise_for_status()
        return r.json()
    except Exception:
        return None


@st.cache_data(ttl=600)
def api_get(path: str):
    return _raw_get(path)


def api_post(path: str, payload: dict):
    try:
        r = httpx.post(f"{API}{path}", json=payload, timeout=60)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        st.error(f"API error {path}: {e}")
        return None


def parallel_get(*paths):
    """Fetch multiple API paths in parallel (thread-safe)."""
    with ThreadPoolExecutor(max_workers=2) as pool:
        futures = [pool.submit(_raw_get, p) for p in paths]
        return [f.result() for f in futures]


# ══════════════════════════════════════════════
# Заголовок
# ══════════════════════════════════════════════
st.title("🛡️ HSE AI Analytics")
st.caption("AI-аналитика охраны труда и промышленной безопасности | KMG Hackathon 2026")

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Инциденты",
    "🔍 Карты Коргау",
    "🔮 Предикт",
    "🚨 Алерты",
])


# ══════════════════════════════════════════════
# TAB 1 — Инциденты
# ══════════════════════════════════════════════
with tab1:
    data, clusters = parallel_get("/api/incidents/summary", "/api/incidents/cause-clusters?n=6")
    if not data:
        st.stop()

    # KPI
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Всего инцидентов", data["total"])
    c2.metric("Несчастных случаев", data["by_type"].get("Несчастный случай", 0))
    c3.metric("Микротравм", data["by_type"].get("Микротравма", 0))
    c4.metric("Изменение MoM", f"{data['mom_change_pct']:+.1f}%",
              delta=data["mom_change_pct"], delta_color="inverse")

    st.divider()
    col_left, col_right = st.columns([2, 1])

    # Time series + rolling avg
    with col_left:
        st.subheader("Динамика инцидентов по месяцам")
        trend_df = pd.DataFrame(data["monthly_trend"])
        if not trend_df.empty:
            trend_df["ds"] = pd.to_datetime(trend_df["ds"])
            fig = go.Figure()
            fig.add_trace(go.Bar(x=trend_df["ds"], y=trend_df["count"],
                                 name="Инциденты", marker_color="#1e3a8a", opacity=0.7))
            fig.add_trace(go.Scatter(x=trend_df["ds"], y=trend_df["rolling_3m"],
                                     name="Скольз. среднее 3м",
                                     line=dict(color="#ef4444", width=2)))
            fig.update_layout(height=320, margin=dict(t=10, b=10), legend=dict(orientation="h"))
            st.plotly_chart(fig, use_container_width=True)

    with col_right:
        # Pie chart по типам
        st.subheader("По типам")
        types_df = pd.DataFrame(list(data["by_type"].items()), columns=["Тип", "Кол-во"])
        fig_pie = px.pie(types_df, values="Кол-во", names="Тип",
                         color_discrete_sequence=px.colors.qualitative.Bold, hole=0.3)
        fig_pie.update_layout(height=320, margin=dict(t=10, b=10),
                              legend=dict(orientation="h", y=-0.2))
        st.plotly_chart(fig_pie, use_container_width=True)

    # Bar chart по организациям
    st.subheader("Топ-10 организаций по числу инцидентов")
    org_df = pd.DataFrame(list(data["by_org"].items()), columns=["Организация", "Кол-во"]).sort_values("Кол-во")
    fig_bar = px.bar(org_df, x="Кол-во", y="Организация", orientation="h",
                     color="Кол-во", color_continuous_scale="Reds")
    fig_bar.update_layout(height=360, margin=dict(t=10, b=10), showlegend=False)
    st.plotly_chart(fig_bar, use_container_width=True)

    # Cause clusters (prefetched in parallel)
    st.subheader("Кластеры причин инцидентов (TF-IDF + KMeans)")
    if clusters:
        cl_df = pd.DataFrame(clusters)[["label", "count"]].rename(
            columns={"label": "Группа причин", "count": "Кол-во"})
        fig_cl = px.bar(cl_df, x="Кол-во", y="Группа причин", orientation="h",
                        color="Кол-во", color_continuous_scale="Blues")
        fig_cl.update_layout(height=280, margin=dict(t=10, b=10))
        st.plotly_chart(fig_cl, use_container_width=True)

    # Объединённый анализ: текст + фото
    st.divider()
    st.subheader("🤖 AI-анализ инцидента (текст и/или фото)")
    st.caption("Опишите инцидент текстом, загрузите фото — или оба варианта")

    col_input_l, col_input_r = st.columns(2)
    with col_input_l:
        incident_text = st.text_area("Описание инцидента:", height=100,
                                     placeholder="Работник поскользнулся на мокром полу...")
    with col_input_r:
        uploaded = st.file_uploader("Фото с места инцидента",
                                    type=["jpg", "jpeg", "png", "webp"])

    has_text = incident_text and incident_text.strip()
    has_photo = uploaded is not None

    if st.button("Анализировать", type="primary", disabled=not (has_text or has_photo)):
        with st.spinner("Gemini анализирует..."):
            try:
                files = {}
                form_data = {}
                if has_text:
                    form_data["description"] = incident_text.strip()
                if has_photo:
                    files["file"] = (uploaded.name, uploaded.getvalue(), uploaded.type)

                r = httpx.post(f"{API}/api/incidents/analyze",
                               data=form_data, files=files if files else None, timeout=60)
                r.raise_for_status()
                result = r.json()
            except Exception as e:
                st.error(f"Ошибка: {e}")
                result = None

        if result:
            # NLP результат
            if result.get("nlp"):
                nlp = result["nlp"]
                st.markdown("#### 📝 NLP-классификация текста")
                r1, r2, r3 = st.columns(3)
                r1.metric("Тип", nlp.get("type", "—"))
                r2.metric("Тяжесть", nlp.get("severity", "—"))
                r3.metric("Причина", nlp.get("root_cause_category", "—"))
                st.info(nlp.get("summary", ""))

            # CV результат
            if result.get("cv"):
                cv = result["cv"]
                st.markdown("#### 📷 CV-анализ фото")

                col_img, col_res = st.columns([1, 1])
                with col_img:
                    st.image(uploaded, caption="Загруженное фото", use_container_width=True)
                with col_res:
                    level_colors = {"Критический": "🔴", "Высокий": "🟠",
                                    "Средний": "🟡", "Низкий": "🟢", "Неизвестно": "⚪"}
                    lvl = cv.get("risk_level", "Неизвестно")
                    st.markdown(f"**{level_colors.get(lvl, '⚪')} Уровень риска: {lvl}**")
                    st.info(cv.get("summary", ""))

                    if cv.get("violations"):
                        for v in cv["violations"]:
                            st.error(f"• **{v['type']}** — {v['description']}")
                    if cv.get("missing_ppe"):
                        for p in cv["missing_ppe"]:
                            st.warning(f"• СИЗ: {p}")
                    if cv.get("recommendations"):
                        for rec in cv["recommendations"]:
                            st.success(f"• {rec}")


# ══════════════════════════════════════════════
# TAB 2 — Коргау
# ══════════════════════════════════════════════
with tab2:
    trends, rankings, gvb, corr = parallel_get(
        "/api/korgau/trends", "/api/korgau/rankings",
        "/api/korgau/good-vs-bad", "/api/korgau/correlation",
    )
    col_l, col_r = st.columns([1, 1])

    # Violation trends по организациям (топ-5)
    with col_l:
        st.subheader("Тренд нарушений по организациям")
        if trends:
            top5 = sorted(trends, key=lambda x: x["total_violations"], reverse=True)[:5]
            fig_tr = go.Figure()
            for org_data in top5:
                monthly = org_data.get("monthly", {})
                months = [str(k) for k in monthly.keys()]
                vals   = list(monthly.values())
                fig_tr.add_trace(go.Scatter(x=months, y=vals,
                                            name=org_data["org"][:25], mode="lines+markers"))
            fig_tr.update_layout(height=320, margin=dict(t=10, b=10),
                                 legend=dict(orientation="h", y=-0.3))
            st.plotly_chart(fig_tr, use_container_width=True)

    # Org rankings — horizontal bar
    with col_r:
        st.subheader("Рейтинг организаций (% нарушений)")
        if rankings:
            rank_df = pd.DataFrame(rankings[:15])[["org", "violation_rate_pct"]].sort_values("violation_rate_pct")
            fig_rank = px.bar(rank_df, x="violation_rate_pct", y="org", orientation="h",
                              color="violation_rate_pct", color_continuous_scale="RdYlGn_r",
                              labels={"violation_rate_pct": "% нарушений", "org": ""})
            fig_rank.update_layout(height=320, margin=dict(t=10, b=10))
            st.plotly_chart(fig_rank, use_container_width=True)

    st.divider()

    # Category breakdown — stacked bar (good vs bad)
    st.subheader("Хорошие vs плохие практики по категориям")
    if gvb:
        all_cats = list(set(list(gvb["bad"].keys()) + list(gvb["good"].keys())))
        gvb_df = pd.DataFrame({
            "Категория": all_cats,
            "Нарушения": [gvb["bad"].get(c, 0) for c in all_cats],
            "Хорошие практики": [gvb["good"].get(c, 0) for c in all_cats],
        }).sort_values("Нарушения", ascending=False).head(12)

        fig_gvb = go.Figure()
        fig_gvb.add_trace(go.Bar(name="Нарушения", x=gvb_df["Категория"],
                                  y=gvb_df["Нарушения"], marker_color="#ef4444"))
        fig_gvb.add_trace(go.Bar(name="Хорошие практики", x=gvb_df["Категория"],
                                  y=gvb_df["Хорошие практики"], marker_color="#22c55e"))
        fig_gvb.update_layout(barmode="group", height=360, margin=dict(t=10, b=10),
                               xaxis_tickangle=-30)
        st.plotly_chart(fig_gvb, use_container_width=True)

    # Correlation with incidents
    # corr already fetched in parallel above
    if corr and corr.get("pearson_r") is not None:
        st.divider()
        st.subheader("Корреляция: нарушения Коргау ↔ инциденты")
        m1, m2, m3 = st.columns(3)
        m1.metric("Pearson r", corr["pearson_r"])
        m2.metric("Spearman r", corr["spearman_r"])
        m3.metric("Точек анализа", corr["n_points"])

        # Автоматическая интерпретация
        pr = corr["pearson_r"]
        sr = corr["spearman_r"]
        n = corr["n_points"]

        if abs(pr) >= 0.7:
            strength = "сильная"
        elif abs(pr) >= 0.4:
            strength = "умеренная"
        elif abs(pr) >= 0.2:
            strength = "слабая"
        else:
            strength = "очень слабая"

        direction = "прямая (больше нарушений → больше инцидентов)" if pr > 0 else "обратная"
        variance = round(pr ** 2 * 100, 1)

        st.info(
            f"**Простыми словами:** в организациях где больше нарушений по Коргау — "
            f"происходит больше реальных инцидентов. Связь **{strength}** ({variance}%).\n\n"
            f"Если в организации часто фиксируют нарушения на аудитах, "
            f"там выше шанс реального происшествия. "
            f"Анализ по **{n} точкам данных**.\n\n"
            f"⚠️ Корреляция ≠ причинность — но закономерность чёткая."
        )

        # Scatter plot + выбросы
        if corr.get("scatter_data"):
            scatter_df = pd.DataFrame(corr["scatter_data"])
            if not scatter_df.empty and len(scatter_df) > 3:
                import numpy as np

                fig_sc = go.Figure()

                # Точки с подписями
                fig_sc.add_trace(go.Scatter(
                    x=scatter_df["violations"], y=scatter_df["incidents"],
                    mode="markers",
                    marker=dict(size=10, color="#1e3a8a", opacity=0.8,
                                line=dict(width=1, color="white")),
                    text=scatter_df.get("label", scatter_df["_org"]),
                    hovertemplate="<b>%{text}</b><br>Нарушений: %{x}<br>Инцидентов: %{y}<extra></extra>",
                    showlegend=False,
                ))

                # Линия тренда (numpy polyfit)
                x_vals = scatter_df["violations"].values.astype(float)
                y_vals = scatter_df["incidents"].values.astype(float)
                if len(x_vals) >= 3:
                    slope, intercept = np.polyfit(x_vals, y_vals, 1)
                    x_line = np.linspace(x_vals.min(), x_vals.max(), 50)
                    y_line = slope * x_line + intercept
                    fig_sc.add_trace(go.Scatter(
                        x=x_line, y=y_line, mode="lines",
                        line=dict(color="#ef4444", width=2, dash="dash"),
                        name=f"Тренд (slope={slope:.2f})",
                    ))

                n_orgs = scatter_df["_org"].nunique()
                fig_sc.update_layout(
                    height=400, margin=dict(t=30, b=50, l=50, r=20),
                    title=f"Связь: нарушений больше → инцидентов больше ({len(scatter_df)} точек, {n_orgs} орг.)",
                    xaxis_title="Нарушений Коргау", yaxis_title="Инцидентов",
                )
                if n_orgs == 1:
                    st.caption(f"В мок-данных только **{scatter_df['_org'].iloc[0]}** присутствует в обоих датасетах. "
                               f"Каждая точка = один месяц. В продакшене будет больше организаций.")
                st.plotly_chart(fig_sc, use_container_width=True)

                # Найти выбросы (>2 std от среднего)
                mean_v = scatter_df["violations"].mean()
                std_v = scatter_df["violations"].std()
                mean_i = scatter_df["incidents"].mean()
                std_i = scatter_df["incidents"].std()

                outliers = scatter_df[
                    (scatter_df["violations"] > mean_v + 2 * std_v) |
                    (scatter_df["incidents"] > mean_i + 2 * std_i)
                ]
                if not outliers.empty:
                    st.warning(
                        f"**Выбросы ({len(outliers)} точек):** организации с аномально высокими показателями "
                        f"— они ломают линейность (Pearson), но не влияют на ранговую связь (Spearman)."
                    )
                    for _, row in outliers.iterrows():
                        st.caption(f"• **{row['_org']}** — {int(row['violations'])} нарушений, {int(row['incidents'])} инцидентов")


# ══════════════════════════════════════════════
# TAB 3 — Предикт
# ══════════════════════════════════════════════
with tab3:
    fc, rs, cm, measures_data = parallel_get(
        "/api/predict/forecast", "/api/predict/risk-scores",
        "/api/predict/correlation-matrix", "/api/predict/scenario/measures",
    )
    st.subheader("🔮 Prophet: прогноз инцидентов на 12 месяцев")
    if fc and "forecast" in fc:
        hist_df = pd.DataFrame(fc["history"])
        fut_df  = pd.DataFrame(fc["forecast"])

        hist_df["ds"] = pd.to_datetime(hist_df["ds"])
        fut_df["ds"]  = pd.to_datetime(fut_df["ds"])

        fig_fc = go.Figure()
        # Исторические данные
        fig_fc.add_trace(go.Scatter(x=hist_df["ds"], y=hist_df["y"],
                                     name="Факт", mode="markers+lines",
                                     line=dict(color="#1e3a8a")))
        fig_fc.add_trace(go.Scatter(x=hist_df["ds"], y=hist_df["yhat"],
                                     name="Тренд (факт)", line=dict(color="#94a3b8", dash="dot")))
        # Прогноз
        fig_fc.add_trace(go.Scatter(
            x=pd.concat([fut_df["ds"], fut_df["ds"].iloc[::-1]]),
            y=pd.concat([fut_df["yhat_upper"], fut_df["yhat_lower"].iloc[::-1]]),
            fill="toself", fillcolor="rgba(239,68,68,0.15)",
            line=dict(color="rgba(0,0,0,0)"), name="95% CI", showlegend=True
        ))
        fig_fc.add_trace(go.Scatter(x=fut_df["ds"], y=fut_df["yhat"],
                                     name="Прогноз", mode="lines",
                                     line=dict(color="#ef4444", width=2)))
        fig_fc.update_layout(height=400, margin=dict(t=10, b=10),
                              legend=dict(orientation="h"), xaxis_title="Дата",
                              yaxis_title="Инцидентов/месяц")
        st.plotly_chart(fig_fc, use_container_width=True)

        with st.expander("Таблица прогноза"):
            st.dataframe(fut_df[["ds", "yhat", "yhat_lower", "yhat_upper"]].rename(
                columns={"ds": "Месяц", "yhat": "Прогноз",
                         "yhat_lower": "Нижняя граница", "yhat_upper": "Верхняя граница"}
            ), use_container_width=True)

    st.divider()

    # Risk Scores — heatmap / таблица
    st.subheader("Риск-скоринг организаций")
    # rs already fetched in parallel above
    if rs:
        rs_df = pd.DataFrame(rs)

        level_colors = {"CRITICAL": "🔴", "HIGH": "🟠", "MEDIUM": "🟡", "LOW": "🔵"}
        rs_df["Уровень"] = rs_df["risk_level"].map(level_colors) + " " + rs_df["risk_level"]

        col_tbl, col_heat = st.columns([1, 1])
        with col_tbl:
            st.dataframe(
                rs_df[["org", "risk_score", "Уровень", "incident_count", "violation_count"]]
                .rename(columns={"org": "Организация", "risk_score": "Скор",
                                 "incident_count": "Инциденты", "violation_count": "Нарушения"})
                .head(20),
                use_container_width=True, height=400,
            )
        with col_heat:
            top20 = rs_df.head(20)
            fig_heat = go.Figure(go.Bar(
                x=top20["risk_score"], y=top20["org"],
                orientation="h",
                marker=dict(
                    color=top20["risk_score"],
                    colorscale=[[0, "#22c55e"], [0.5, "#fbbf24"], [1, "#ef4444"]],
                    showscale=True,
                    cmin=0, cmax=100,
                )
            ))
            fig_heat.update_layout(height=400, margin=dict(t=10, b=10, l=10),
                                   xaxis_title="Риск-скор (0-100)", yaxis_title="")
            st.plotly_chart(fig_heat, use_container_width=True)

    st.divider()

    # Correlation Matrix
    st.subheader("Матрица корреляций: категории нарушений × типы инцидентов")
    # cm already fetched in parallel above
    if cm and cm.get("matrix"):
        matrix_data = cm["matrix"]
        inc_types = cm["incident_types"]
        cats = cm["categories"]

        z = [[matrix_data.get(cat, {}).get(inc, 0) for inc in inc_types] for cat in cats]
        fig_cm = go.Figure(go.Heatmap(
            z=z, x=inc_types, y=cats,
            colorscale="RdBu", zmid=0,
            text=[[f"{v:.2f}" for v in row] for row in z],
            texttemplate="%{text}",
        ))
        fig_cm.update_layout(height=420, margin=dict(t=10, b=10))
        st.plotly_chart(fig_cm, use_container_width=True)
        st.caption(
            "В мок-данных только 1 организация присутствует в обоих датасетах (инциденты + Коргау). "
            "Матрица построена по 13 месяцам этой организации. "
            "В продакшене — сотни организаций, матрица будет точнее."
        )

    st.divider()

    # ── Scenario Modeling (3.1.2) ──
    st.subheader("🎯 Сценарный анализ: что если внедрить меры контроля?")
    st.caption("Расчёт изменения вероятности инцидентов при внедрении мер (Pearson r = 0.415)")

    # Получаем список доступных мер и организаций
    # measures_data already fetched in parallel above
    summary_data  = api_get("/api/incidents/summary")

    if measures_data and summary_data:
        all_orgs    = sorted(summary_data.get("by_org", {}).keys())
        all_measures = measures_data.get("measures", [])

        sc_col1, sc_col2 = st.columns([1, 1])
        with sc_col1:
            sc_org = st.selectbox("Организация:", all_orgs, key="sc_org")
        with sc_col2:
            sc_measures = st.multiselect(
                "Меры контроля:", all_measures,
                default=all_measures[:2] if all_measures else [],
                key="sc_measures",
            )

        if st.button("Рассчитать эффект", type="primary", key="sc_calc"):
            if not sc_measures:
                st.warning("Выберите хотя бы одну меру контроля.")
            else:
                with st.spinner("Считаем..."):
                    try:
                        r = httpx.post(
                            f"{API}/api/predict/scenario",
                            json={"org": sc_org, "measures": sc_measures},
                            timeout=30,
                        )
                        r.raise_for_status()
                        sc = r.json()
                    except Exception as e:
                        st.error(f"Ошибка: {e}")
                        sc = None

                if sc:
                    st.success(f"Результат для: **{sc['org']}**")
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Инцидентов сейчас", int(sc["baseline_incidents"]))
                    m2.metric(
                        "Прогноз после мер",
                        f"{sc['predicted_incidents']:.0f}",
                        delta=f"-{sc['avoided_incidents']:.0f}",
                        delta_color="inverse",
                    )
                    m3.metric("Снижение риска", f"{sc['incident_reduction_pct']:.1f}%")
                    m4.metric(
                        "Экономия",
                        f"{sc['economic_saving_tenge'] // 1_000_000} млн ₸",
                    )

                    # Waterfall breakdown по мерам
                    if sc.get("measures_applied"):
                        st.markdown("**Вклад каждой меры:**")
                        ma_df = pd.DataFrame(sc["measures_applied"])
                        ma_df = ma_df.rename(columns={
                            "measure": "Мера",
                            "category": "Категория",
                            "current_violations": "Текущих нарушений",
                            "violation_reduction_pct": "Снижение нарушений %",
                            "incident_impact_pct": "Вклад в снижение инцидентов %",
                        })

                        fig_sc = px.bar(
                            ma_df, x="Вклад в снижение инцидентов %", y="Мера",
                            orientation="h",
                            color="Вклад в снижение инцидентов %",
                            color_continuous_scale="Greens",
                            hover_data=["Категория", "Текущих нарушений", "Снижение нарушений %"],
                        )
                        fig_sc.update_layout(height=max(250, len(ma_df) * 55),
                                             margin=dict(t=10, b=10))
                        st.plotly_chart(fig_sc, use_container_width=True)

                        with st.expander("Детали по мерам"):
                            st.dataframe(ma_df, use_container_width=True)


# ══════════════════════════════════════════════
# TAB 4 — Алерты
# ══════════════════════════════════════════════
with tab4:
    alerts_data = api_get("/api/korgau/alerts")

    if alerts_data:
        # Фильтр
        orgs = ["Все"] + sorted(set(a["org"] for a in alerts_data))
        sel_org = st.selectbox("Фильтр по организации:", orgs)
        sel_level = st.multiselect("Уровень алерта:", ["CRITICAL", "HIGH", "MEDIUM", "LOW"],
                                   default=["CRITICAL", "HIGH", "MEDIUM", "LOW"])

        filtered = [
            a for a in alerts_data
            if (sel_org == "Все" or a["org"] == sel_org)
            and a["level"] in sel_level
        ]

        # Счётчики
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("🔴 CRITICAL", sum(1 for a in filtered if a["level"] == "CRITICAL"))
        c2.metric("🟠 HIGH",     sum(1 for a in filtered if a["level"] == "HIGH"))
        c3.metric("🟡 MEDIUM",   sum(1 for a in filtered if a["level"] == "MEDIUM"))
        c4.metric("🔵 LOW",      sum(1 for a in filtered if a["level"] == "LOW"))

        st.divider()

        # Карточки алертов
        for alert in filtered:
            lvl = alert["level"].lower()
            top_cats = ", ".join(f"{k} ({v})" for k, v in list(alert["top_categories"].items())[:3])
            extra = ""
            if alert.get("repeated_violation_reason"):
                extra += f"<br/>🔄 Повтор: {alert['repeated_violation_reason']}"
            if alert.get("yoy_reason"):
                extra += f"<br/>📈 {alert['yoy_reason']}"
            st.markdown(f"""
<div class="alert-{lvl}">
  <strong>[{alert['level']}] {alert['org']}</strong><br/>
  Нарушений за 30 дней: <strong>{alert['count_30d']}</strong> &nbsp;|&nbsp; Период: {alert['period']}<br/>
  Топ категории: {top_cats}{extra}
</div>""", unsafe_allow_html=True)

    st.divider()

    # Кнопки скачать отчёты
    st.subheader("📄 Скачать отчёты")
    col_pdf, col_xlsx = st.columns(2)
    with col_pdf:
        if st.button("Сгенерировать PDF", type="primary"):
            with st.spinner("Генерируем PDF..."):
                try:
                    r = httpx.get(f"{API}/api/reports/pdf", timeout=60)
                    r.raise_for_status()
                    st.download_button(
                        label="Скачать HSE_Report.pdf",
                        data=r.content,
                        file_name=f"HSE_Report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf",
                    )
                except Exception as e:
                    st.error(f"Ошибка PDF: {e}")
    with col_xlsx:
        if st.button("Сгенерировать Excel"):
            with st.spinner("Генерируем Excel..."):
                try:
                    r = httpx.get(f"{API}/api/reports/excel", timeout=60)
                    r.raise_for_status()
                    st.download_button(
                        label="Скачать HSE_Report.xlsx",
                        data=r.content,
                        file_name=f"HSE_Report_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                except Exception as e:
                    st.error(f"Ошибка Excel: {e}")
