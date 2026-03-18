import streamlit as st
import httpx
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd

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


@st.cache_data(ttl=300)
def api_get(path: str):
    try:
        r = httpx.get(f"{API}{path}", timeout=30)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        st.error(f"API error {path}: {e}")
        return None


@st.cache_data(ttl=300)
def api_post(path: str, payload: dict):
    try:
        r = httpx.post(f"{API}{path}", json=payload, timeout=30)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        st.error(f"API error {path}: {e}")
        return None


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
    data = api_get("/api/incidents/summary")
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

    # Cause clusters
    st.subheader("Кластеры причин инцидентов (TF-IDF + KMeans)")
    clusters = api_get("/api/incidents/cause-clusters?n=6")
    if clusters:
        cl_df = pd.DataFrame(clusters)[["label", "count"]].rename(
            columns={"label": "Группа причин", "count": "Кол-во"})
        fig_cl = px.bar(cl_df, x="Кол-во", y="Группа причин", orientation="h",
                        color="Кол-во", color_continuous_scale="Blues")
        fig_cl.update_layout(height=280, margin=dict(t=10, b=10))
        st.plotly_chart(fig_cl, use_container_width=True)

    # NLP классификатор
    st.divider()
    st.subheader("🤖 NLP-классификатор инцидентов (Gemini)")
    txt = st.text_area("Описание инцидента:", height=80,
                       placeholder="Работник поскользнулся на мокром полу...")
    if st.button("Классифицировать", type="primary"):
        if txt.strip():
            with st.spinner("Gemini анализирует..."):
                result = api_post("/api/incidents/classify", {"description": txt})
            if result:
                r1, r2, r3 = st.columns(3)
                r1.metric("Тип", result["type"])
                r2.metric("Тяжесть", result["severity"])
                r3.metric("Корневая причина", result["root_cause_category"])
                st.info(result["summary"])

    # CV-анализ фото (F-08)
    st.divider()
    st.subheader("📷 CV-анализ фото с места происшествия (Gemini Vision)")
    uploaded = st.file_uploader(
        "Загрузить фото с места инцидента",
        type=["jpg", "jpeg", "png", "webp"],
        help="Gemini проанализирует фото и выявит нарушения, отсутствующие СИЗ и опасные условия"
    )
    if uploaded and st.button("Анализировать фото", type="primary"):
        col_img, col_res = st.columns([1, 1])
        with col_img:
            st.image(uploaded, caption="Загруженное фото", use_container_width=True)
        with col_res:
            with st.spinner("Gemini Vision анализирует фото..."):
                try:
                    r = httpx.post(
                        f"{API}/api/incidents/analyze-photo",
                        files={"file": (uploaded.name, uploaded.getvalue(), uploaded.type)},
                        timeout=60,
                    )
                    r.raise_for_status()
                    cv = r.json()
                except Exception as e:
                    st.error(f"Ошибка: {e}")
                    cv = None

            if cv:
                level_colors = {
                    "Критический": "🔴", "Высокий": "🟠",
                    "Средний": "🟡", "Низкий": "🟢", "Неизвестно": "⚪"
                }
                lvl = cv.get("risk_level", "Неизвестно")
                st.markdown(f"### {level_colors.get(lvl, '⚪')} Уровень риска: **{lvl}**")
                st.info(cv.get("summary", ""))

                if cv.get("violations"):
                    st.markdown("**Выявленные нарушения:**")
                    for v in cv["violations"]:
                        st.error(f"• **{v['type']}** — {v['description']}")

                if cv.get("missing_ppe"):
                    st.markdown("**Отсутствующие СИЗ:**")
                    for p in cv["missing_ppe"]:
                        st.warning(f"• {p}")

                if cv.get("unsafe_conditions"):
                    st.markdown("**Опасные условия:**")
                    for c in cv["unsafe_conditions"]:
                        st.warning(f"• {c}")

                if cv.get("recommendations"):
                    st.markdown("**Рекомендации:**")
                    for rec in cv["recommendations"]:
                        st.success(f"• {rec}")


# ══════════════════════════════════════════════
# TAB 2 — Коргау
# ══════════════════════════════════════════════
with tab2:
    col_l, col_r = st.columns([1, 1])

    # Violation trends по организациям (топ-5)
    with col_l:
        st.subheader("Тренд нарушений по организациям")
        trends = api_get("/api/korgau/trends")
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
        rankings = api_get("/api/korgau/rankings")
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
    gvb = api_get("/api/korgau/good-vs-bad")
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
    corr = api_get("/api/korgau/correlation")
    if corr and corr.get("pearson_r") is not None:
        st.divider()
        st.subheader("Корреляция: нарушения Коргау ↔ инциденты")
        m1, m2, m3 = st.columns(3)
        m1.metric("Pearson r", corr["pearson_r"])
        m2.metric("Spearman r", corr["spearman_r"])
        m3.metric("Точек анализа", corr["n_points"])


# ══════════════════════════════════════════════
# TAB 3 — Предикт
# ══════════════════════════════════════════════
with tab3:
    st.subheader("🔮 Prophet: прогноз инцидентов на 12 месяцев")
    fc = api_get("/api/predict/forecast")
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
    rs = api_get("/api/predict/risk-scores")
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
    cm = api_get("/api/predict/correlation-matrix")
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

    st.divider()

    # ── Scenario Modeling (3.1.2) ──
    st.subheader("🎯 Сценарный анализ: что если внедрить меры контроля?")
    st.caption("Расчёт изменения вероятности инцидентов при внедрении мер (Pearson r = 0.415)")

    # Получаем список доступных мер и организаций
    measures_data = api_get("/api/predict/scenario/measures")
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
            st.markdown(f"""
<div class="alert-{lvl}">
  <strong>[{alert['level']}] {alert['org']}</strong><br/>
  Нарушений за 30 дней: <strong>{alert['count_30d']}</strong> &nbsp;|&nbsp; Период: {alert['period']}<br/>
  Топ категории: {top_cats}
</div>""", unsafe_allow_html=True)

    st.divider()

    # Кнопка скачать PDF
    st.subheader("📄 Скачать полный PDF-отчёт")
    if st.button("Сгенерировать и скачать PDF", type="primary"):
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
                st.error(f"Ошибка генерации PDF: {e}")
