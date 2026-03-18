
# 🛡️ HSE AI Analytics — Охрана труда

> **KMG Hackathon 2026** | AI-аналитика охраны труда и промышленной безопасности

AI-аналитический слой поверх двух модулей HSE-системы, который переводит исторические данные о происшествиях и наблюдениях в **предиктивный инструмент управления безопасностью** с измеримым экономическим эффектом.

---

## Покрытие требований ТЗ

### Компонент A — AI-аналитика происшествий

| # | Требование | Приоритет | Реализация |
|---|-----------|-----------|-----------|
| F-01 | NLP-классификация происшествий по типу | Высокий | ✅ `nlp_service.py` → Gemini `gemini-2.0-flash`, `POST /api/incidents/classify` |
| F-02 | Паттерны по типу, организации, локации, времени суток/дня недели | Высокий | ✅ `incident_analyzer.get_summary()` → by_type, by_org, by_location, `_shift`, `_weekday` |
| F-03 | Кластеризация причин инцидентов | Высокий | ✅ `incident_analyzer.get_cause_clusters()` → TF-IDF (bigrams) + KMeans 6 кластеров |
| F-04 | Генерация рекомендаций по мерам контроля | Высокий | ✅ `nlp_service.generate_recommendations()` → Gemini топ-5, `GET /api/incidents/recommendations` |
| F-05 | Сводный дашборд: период, тип, организация, локация | Высокий | ✅ Streamlit Tab 1 — time series, pie, bar, таблица |
| F-06 | Предиктив на 12 / 6 / 3 месяца с доверительным интервалом | Высокий | ✅ `predictor.get_forecast()` → Prophet + 95% CI, `GET /api/predict/forecast?periods=12` |
| F-06b | Сценарный анализ мер контроля (3.1.2) | Высокий | ✅ `predictor.calculate_scenario()` → 9 мер, Pearson r=0.415, `POST /api/predict/scenario` |
| F-07 | Расчёт экономического эффекта | Средний | ✅ KPI-карточки в PDF и Streamlit: ~7 НС, ~48 микротравм, ~121 млн ₸/год |
| F-08 | CV-анализ фото с места происшествий | Средний | ✅ `nlp_service.analyze_incident_photo()` → Gemini Vision, `POST /api/incidents/analyze-photo` |
| F-09 | Экспорт отчётов в PDF | Низкий | ✅ `report_generator.py` → fpdf2, 4 страницы, `GET /api/reports/pdf` |

### Компонент B — AI-аналитика Карт Коргау

| # | Требование | Приоритет | Реализация |
|---|-----------|-----------|-----------|
| F-10 | Паттерны нарушений по организациям, категориям, периодам | Высокий | ✅ `korgau_analyzer.get_violation_rates()` → groupby (org, category, month) |
| F-11 | Система алертов при превышении пороговой частоты | Высокий | ✅ `korgau_analyzer.get_alerts()` → 4 уровня, `GET /api/korgau/alerts`, Streamlit Tab 4 |
| F-12 | Рейтинг организаций по соответствию требованиям ОТ | Высокий | ✅ `korgau_analyzer.get_org_rankings()` → 64 организации, `GET /api/korgau/rankings` |
| F-13 | Корреляция нарушений Коргау с реальными инцидентами | Высокий | ✅ `korgau_analyzer.get_correlation_with_incidents()` → Pearson r=0.415, Spearman r=0.501 |
| F-14 | Pre-alert при росте нарушений в зонах риска | Высокий | ✅ `korgau_analyzer.get_org_trends()` → direction: рост/снижение/стабильно |
| F-15 | AI-классификация наблюдений по кластерам (СИЗ, LOTO и др.) | Средний | ✅ `nlp_service.classify_korgau_parsed()` → Gemini, 10 тематических кластеров |

### MVP-чеклист

| MVP-функция | Статус | Реализация |
|-------------|--------|-----------|
| Дашборд статистики происшествий | ✅ | Streamlit Tab 1 — фильтры, графики динамики |
| Предиктив на 12 / 6 / 3 месяца | ✅ | Prophet forecast с CI-лентой, Streamlit Tab 3 |
| Топ-5 зон риска | ✅ | Risk Scoring 0–100, heatmap по организациям |
| Алерты по Карте Коргау | ✅ | 4 уровня (CRITICAL/HIGH/MEDIUM/LOW), цветные карточки |
| Рекомендации от AI | ✅ | Топ-5 от Gemini с приоритетом и обоснованием |
| Расчёт экономического эффекта | ✅ | KPI-карточки: люди + тенге в PDF и дашборде |
| Telegram-алерты | ⭕ | Опционально, не реализовано |

---

## Архитектура (5.1 Поток данных)

```
incidents.xlsx   ─┐
                  ├─► data_loader.py (ETL) ─► pandas DataFrame (Data Warehouse)
korgau_cards.xlsx ─┘         │
                              ▼
                   ┌─────────────────────┐
                   │      ML Core        │
                   │  Prophet (forecast) │
                   │  TF-IDF + KMeans    │
                   │  Risk Scoring       │
                   │  Pearson/Spearman   │
                   └─────────┬───────────┘
                             │
                   ┌─────────▼───────────┐
                   │     NLP Engine      │
                   │  Gemini 3 Flash     │
                   │  классификация      │
                   │  рекомендации       │
                   └─────────┬───────────┘
                             │
                   ┌─────────▼───────────┐
                   │    FastAPI (8002)    │
                   │    15 эндпоинтов    │
                   └──┬──────┬──────┬────┘
                      │      │      │
               Dashboard  Алерты  PDF
             (Streamlit)         (fpdf2)
```

### Трёхуровневая архитектура

| Уровень | Компонент | Реализация |
|---------|-----------|-----------|
| **Данные** | ETL Pipeline + Data Warehouse | `data_loader.py` — Extract (xlsx) → Transform (очистка, смены, стаж, флаги) → Load (pandas in-memory, `@lru_cache`) |
| **AI** | ML Core + NLP Engine | Prophet, TF-IDF+KMeans, Risk Scoring, Pearson/Spearman + Gemini `gemini-2.0-flash` |
| **Представление** | Dashboard + Alert Manager + Report Generator | Streamlit 4 таба + FastAPI алерты + fpdf2 PDF |

---

## Технический стек

| Компонент | Технология |
|-----------|-----------|
| Backend API | Python 3.11, FastAPI, uvicorn |
| Данные | pandas + openpyxl (in-memory, без БД) |
| ML | Prophet (time series), scikit-learn (TF-IDF + KMeans) |
| NLP | Google Gemini `gemini-2.0-flash` |
| Визуализация | Plotly (10 графиков) |
| Отчёты | fpdf2 — PDF 4 страницы, кириллица (DejaVu TTF) |
| Frontend | Streamlit — 4 таба |
| Конфиг | pydantic-settings + python-dotenv |
| Контейнеры | Docker + Docker Compose |

---

## Данные

| Датасет | Файл | Объём |
|---------|------|-------|
| Происшествия | `data/incidents.xlsx` | 220 строк, 50 колонок |
| Карты Коргау | `data/korgau_cards.xlsx` | 9 917 строк, 16 колонок |

**ETL-обогащение** (сверх сырых данных):

*Инциденты:* `_type` (из флаговых колонок), `_shift` (смена по времени суток), `_weekday`, `_experience_years` (стаж), `_in_work_hours`

*Коргау:* `_is_violation`, `_is_resolved` (статус устранения), `_shift`, `_weekday`, `_reported` (эскалация)

---

## Структура проекта

```
hse/
├── .env                        # API-ключи (не коммитить)
├── .env.example
├── requirements.txt
├── Dockerfile                  # FastAPI + cmdstan (Prophet)
├── Dockerfile.streamlit
├── docker-compose.yml
├── streamlit_app.py            # Frontend — 4 таба, 10 Plotly-графиков
├── app/
│   ├── main.py                 # FastAPI, CORS, роутеры
│   ├── config.py               # pydantic-settings → .env
│   ├── schemas.py              # Pydantic модели
│   ├── api/
│   │   ├── incidents.py        # /api/incidents/*
│   │   ├── korgau.py           # /api/korgau/*
│   │   ├── predict.py          # /api/predict/*
│   │   └── reports.py          # /api/reports/pdf
│   └── services/
│       ├── data_loader.py      # ETL: xlsx → pandas (lru_cache)
│       ├── incident_analyzer.py # groupby, TF-IDF, KMeans, тренды
│       ├── korgau_analyzer.py  # алерты, рейтинг, корреляция
│       ├── nlp_service.py      # Gemini API (lru_cache 512/2048)
│       ├── predictor.py        # Prophet + Risk Scoring + матрица
│       └── report_generator.py # PDF 4 страницы
└── data/
    ├── incidents.xlsx
    └── korgau_cards.xlsx
```

---

## API — 15 эндпоинтов

| Метод | Эндпоинт | Описание |
|-------|----------|----------|
| GET | `/health` | Статус сервера |
| GET | `/api/incidents/summary` | Total, by_type, by_org, тренд, MoM |
| POST | `/api/incidents/classify` | NLP-классификация описания (Gemini) |
| GET | `/api/incidents/cause-clusters` | TF-IDF + KMeans кластеры причин |
| GET | `/api/incidents/org-breakdown` | Детализация + тренд по организациям |
| GET | `/api/incidents/recommendations` | Топ-5 рекомендаций от Gemini |
| GET | `/api/korgau/alerts` | Алерты CRITICAL/HIGH/MEDIUM/LOW за 30 дней |
| GET | `/api/korgau/rankings` | Рейтинг по % нарушений, остановки работ |
| GET | `/api/korgau/trends` | Рост/снижение нарушений по организациям |
| GET | `/api/korgau/good-vs-bad` | Хорошие vs плохие практики по категориям |
| GET | `/api/korgau/correlation` | Pearson/Spearman: нарушения ↔ инциденты |
| GET | `/api/predict/forecast?periods=12` | Prophet прогноз + 95% CI |
| GET | `/api/predict/risk-scores` | Risk Score 0–100 по формуле ТЗ |
| GET | `/api/predict/correlation-matrix` | Матрица: категории нарушений × типы инцидентов |
| GET | `/api/predict/scenario/measures` | Список доступных мер контроля |
| POST | `/api/predict/scenario` | Сценарный анализ: снижение инцидентов при внедрении мер |
| GET | `/api/reports/pdf` | PDF-отчёт 4 страницы |

---

## Предиктивная аналитика

### Prophet Forecast
```
Агрегация по месяцам → Prophet.fit() → forecast(12 мес)
Компоненты: тренд + годовая сезонность + 95% CI
Fallback: линейная регрессия если cmdstan недоступен
```

### Risk Scoring (формула из ТЗ)
```
risk_score = 0.35 × incident_rate_norm
           + 0.30 × violation_rate_norm
           + 0.20 × trend_slope_norm
           + 0.15 × severity_factor_norm

Нормализация: 0–100 по всем организациям
CRITICAL ≥ 70 | HIGH ≥ 50 | MEDIUM ≥ 25 | LOW < 25
```

### Корреляционный анализ
```
Pearson r  = 0.415  (нарушения Коргау → инциденты)
Spearman r = 0.501
```

---

## Система алертов

| Уровень | Порог | Цвет |
|---------|-------|------|
| 🔴 CRITICAL | > 50 нарушений за 30 дней | red |
| 🟠 HIGH | > 30 нарушений за 30 дней | orange |
| 🟡 MEDIUM | > 15 нарушений за 30 дней | yellow |
| 🔵 LOW | > 5 нарушений за 30 дней | blue |

---

## Экономический эффект

### В людях

| Показатель | До AI | После AI | Эффект |
|-----------|-------|----------|--------|
| Несчастных случаев | 18/год | 11/год | ↓ 38% |
| Микротравм | 120/год | 72/год | ↓ 40% |
| Опасных ситуаций | 350/год | 175/год | ↓ 50% |
| Время реагирования | 72 ч | 12 ч | ↓ 83% |
| Охват предиктивными мерами | 0% | до 70% | ↑ |

### В деньгах

| Статья | Экономия/год |
|--------|-------------|
| Прямые затраты на НС (7 × 5 млн ₸) | ≈ 35 000 000 ₸ |
| Косвенные потери (простой, репутация) | ≈ 70 000 000 ₸ |
| Штрафы и предписания (−30%) | ≈ 5 000 000 ₸ |
| Расследования и отчётность (−70%) | ≈ 8 000 000 ₸ |
| ROI автоматизации аудитов (+40%) | ≈ 3 000 000 ₸ |
| **ИТОГО** | **≈ 121 000 000 ₸ (~250 000 USD)** |

---

## Критерии оценки хакатона

| Критерий | Вес | Реализация |
|---------|-----|-----------|
| Точность предиктивной модели | 25% | Prophet + линейный fallback |
| Качество рекомендаций | 20% | Gemini топ-5 с приоритетом и обоснованием |
| UX дашборда | 15% | Streamlit 4 таба, 10 Plotly-графиков |
| Система алертов | 15% | 4 уровня, цветные карточки, фильтры |
| Интеграция в HSE-систему | 15% | 15 REST эндпоинтов, OpenAPI /docs |
| Расчёт экономического эффекта | 10% | KPI-карточки: люди + тенге в PDF и UI |

---

## Быстрый старт

```bash
cp .env.example .env
# Вставить GEMINI_API_KEY в .env

docker compose up --build
```

| Сервис | URL |
|--------|-----|
| Streamlit Dashboard | http://localhost:8502 |
| FastAPI | http://localhost:8002 |
| Swagger UI | http://localhost:8002/docs |

---

> 🎯 **Целевой результат:** рабочий AI-дашборд, который показывает аналитику прошлых инцидентов, предсказывает будущие риски, генерирует рекомендации и показывает экономический эффект — в понятных единицах: **люди и деньги**.
