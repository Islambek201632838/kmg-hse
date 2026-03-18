"""
Генератор презентации HSE AI Analytics — KMG Hackathon 2026.
Запуск: python create_presentation.py
Результат: HSE_Presentation.pptx
"""

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Inches, Pt
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx"])
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import copy

# ── Цветовая палитра ──────────────────────────────────────────────────────────
C_DARK_BLUE  = RGBColor(0x1e, 0x3a, 0x8a)
C_BLUE       = RGBColor(0x2d, 0x6a, 0xd4)
C_RED        = RGBColor(0xef, 0x44, 0x44)
C_AMBER      = RGBColor(0xf5, 0x9e, 0x0b)
C_GREEN      = RGBColor(0x22, 0xc5, 0x5e)
C_WHITE      = RGBColor(0xff, 0xff, 0xff)
C_LIGHT      = RGBColor(0xf1, 0xf5, 0xf9)
C_GRAY       = RGBColor(0x64, 0x74, 0x8b)
C_DARK       = RGBColor(0x1e, 0x29, 0x3b)

W = Inches(13.33)   # широкоформатный слайд 16:9
H = Inches(7.5)


# ── Хелперы ───────────────────────────────────────────────────────────────────

def new_prs() -> Presentation:
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H
    return prs


def blank_slide(prs: Presentation):
    layout = prs.slide_layouts[6]  # полностью пустой
    return prs.slides.add_slide(layout)


def rect(slide, x, y, w, h, fill_rgb: RGBColor, alpha=None):
    shape = slide.shapes.add_shape(1, x, y, w, h)  # MSO_SHAPE_TYPE.RECTANGLE
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    return shape


def txt(slide, text: str, x, y, w, h,
        font_size=18, bold=False, color: RGBColor = C_WHITE,
        align=PP_ALIGN.LEFT, italic=False, wrap=True):
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf  = txb.text_frame
    tf.word_wrap = wrap
    p   = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size   = Pt(font_size)
    run.font.bold   = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb


def add_bullet(tf, text: str, level=0, size=16,
               color: RGBColor = C_DARK, bold=False):
    from pptx.util import Pt
    p = tf.add_paragraph()
    p.level = level
    p.space_before = Pt(4)
    run = p.add_run()
    run.text = text
    run.font.size  = Pt(size)
    run.font.color.rgb = color
    run.font.bold  = bold


def header_bar(slide, title: str, subtitle: str = ""):
    """Синяя полоса сверху с заголовком."""
    rect(slide, 0, 0, W, Inches(1.3), C_DARK_BLUE)
    txt(slide, title, Inches(0.4), Inches(0.15), Inches(10), Inches(0.7),
        font_size=28, bold=True, color=C_WHITE)
    if subtitle:
        txt(slide, subtitle, Inches(0.4), Inches(0.82), Inches(10), Inches(0.4),
            font_size=14, color=RGBColor(0xba, 0xd4, 0xf8))


def kpi_box(slide, value: str, label: str,
            x, y, w=Inches(2.5), h=Inches(1.4),
            bg: RGBColor = C_DARK_BLUE):
    rect(slide, x, y, w, h, bg)
    txt(slide, value, x, y + Inches(0.1), w, Inches(0.7),
        font_size=30, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
    txt(slide, label, x, y + Inches(0.8), w, Inches(0.5),
        font_size=12, color=RGBColor(0xba, 0xd4, 0xf8), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
# СЛАЙДЫ
# ══════════════════════════════════════════════════════════════════════════════

def slide_title(prs):
    """Слайд 1 — Титульный."""
    sl = blank_slide(prs)

    # Фон — тёмно-синий градиент (два прямоугольника)
    rect(sl, 0, 0, W, H, C_DARK_BLUE)
    rect(sl, 0, Inches(5.5), W, Inches(2.0), C_DARK)

    # Декоративная полоса
    rect(sl, 0, Inches(5.3), W, Inches(0.06), C_RED)

    # Логотип-заглушка (щит)
    txt(sl, "🛡️", Inches(0.5), Inches(0.6), Inches(1.5), Inches(1.5),
        font_size=64, color=C_WHITE)

    # Заголовок
    txt(sl, "HSE AI Analytics",
        Inches(2.0), Inches(0.5), Inches(10), Inches(1.3),
        font_size=52, bold=True, color=C_WHITE)

    txt(sl, "AI-аналитика охраны труда и промышленной безопасности",
        Inches(2.0), Inches(1.75), Inches(10), Inches(0.7),
        font_size=22, color=RGBColor(0xba, 0xd4, 0xf8))

    # Разделитель
    rect(sl, Inches(2.0), Inches(2.5), Inches(4.0), Inches(0.05), C_AMBER)

    txt(sl, "KMG Hackathon 2026",
        Inches(2.0), Inches(2.7), Inches(6), Inches(0.5),
        font_size=18, bold=True, color=C_AMBER)

    # Нижняя плашка
    txt(sl, "Предиктивный инструмент управления безопасностью с измеримым экономическим эффектом",
        Inches(0.5), Inches(5.8), Inches(12), Inches(0.6),
        font_size=14, color=RGBColor(0x94, 0xa3, 0xb8), align=PP_ALIGN.CENTER)

    txt(sl, "FastAPI  •  Prophet  •  Gemini 2.0 Flash  •  Streamlit  •  Docker",
        Inches(0.5), Inches(6.4), Inches(12), Inches(0.5),
        font_size=13, color=RGBColor(0x64, 0x74, 0x8b), align=PP_ALIGN.CENTER)


def slide_problem(prs):
    """Слайд 2 — Проблема."""
    sl = blank_slide(prs)
    rect(sl, 0, 0, W, H, C_LIGHT)
    header_bar(sl, "Проблема: реактивный подход к безопасности",
               "Текущее состояние охраны труда в нефтегазовой отрасли")

    # 3 колонки — боль
    cols = [
        ("18\nНС/год",    "Несчастных\nслучаев",       C_RED),
        ("120\nмикротравм", "Ежегодно",                C_AMBER),
        ("72 часа",       "Среднее время\nреагирования", C_BLUE),
    ]
    for i, (val, lbl, color) in enumerate(cols):
        x = Inches(0.5 + i * 4.2)
        kpi_box(sl, val, lbl, x, Inches(1.6), Inches(3.8), Inches(1.6), bg=color)

    # Боли
    pain_y = Inches(3.4)
    rect(sl, Inches(0.4), pain_y, Inches(12.5), Inches(3.6), C_WHITE)

    txb = sl.shapes.add_textbox(Inches(0.7), pain_y + Inches(0.15),
                                Inches(12), Inches(3.2))
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.add_run().text = "Ключевые проблемы"
    p.runs[0].font.size  = Pt(18)
    p.runs[0].font.bold  = True
    p.runs[0].font.color.rgb = C_DARK_BLUE

    items = [
        "❌  Нет предиктивной аналитики — инциденты расследуются постфактум",
        "❌  Данные Карт Коргау и происшествий не связаны между собой",
        "❌  Ручной анализ сотен наблюдений занимает дни",
        "❌  Экономический эффект мер безопасности не рассчитывается",
        "❌  350 опасных ситуаций в год остаются без превентивных мер",
    ]
    for item in items:
        add_bullet(tf, item, size=15, color=C_DARK)

    # Итог
    rect(sl, Inches(0.4), Inches(6.8), Inches(12.5), Inches(0.45), C_RED)
    txt(sl, "Потери: ~200 млн ₸/год  |  Цена бездействия растёт",
        Inches(0.4), Inches(6.8), Inches(12.5), Inches(0.45),
        font_size=15, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)


def slide_solution(prs):
    """Слайд 3 — Решение."""
    sl = blank_slide(prs)
    rect(sl, 0, 0, W, H, C_LIGHT)
    header_bar(sl, "Решение: HSE AI Analytics",
               "AI-аналитический слой поверх двух модулей HSE-системы")

    # Два модуля
    for i, (title, items, color) in enumerate([
        ("Компонент A — Инциденты", [
            "NLP-классификация описаний (Gemini)",
            "Паттерны по типу, орг., локации, смене",
            "Кластеризация причин (TF-IDF + KMeans)",
            "Топ-5 рекомендаций от AI",
            "Prophet-прогноз + сценарный анализ",
            "CV-анализ фото (Gemini Vision)",
        ], C_DARK_BLUE),
        ("Компонент B — Карты Коргау", [
            "Паттерны нарушений по орг./категориям",
            "4-уровневая система алертов",
            "Рейтинг организаций по compliance",
            "Корреляция нарушений с инцидентами",
            "Pre-alert при росте нарушений",
            "AI-кластеризация наблюдений",
        ], C_BLUE),
    ]):
        x = Inches(0.3 + i * 6.6)
        rect(sl, x, Inches(1.45), Inches(6.3), Inches(4.8), color)
        txt(sl, title, x + Inches(0.15), Inches(1.5), Inches(6.0), Inches(0.6),
            font_size=16, bold=True, color=C_WHITE)
        rect(sl, x, Inches(2.1), Inches(6.3), Inches(0.03), C_WHITE)

        txb = sl.shapes.add_textbox(x + Inches(0.15), Inches(2.2),
                                    Inches(6.0), Inches(3.8))
        tf = txb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.add_run().text = ""  # пустой первый параграф
        for item in items:
            add_bullet(tf, f"✓  {item}", size=14, color=C_WHITE)

    # Стрелка объединения
    txt(sl, "+", Inches(6.15), Inches(3.5), Inches(1.0), Inches(0.8),
        font_size=36, bold=True, color=C_DARK_BLUE, align=PP_ALIGN.CENTER)

    # Итог внизу
    rect(sl, Inches(0.3), Inches(6.35), Inches(12.7), Inches(0.9), C_DARK_BLUE)
    txt(sl, "15 REST-эндпоинтов  •  Streamlit-дашборд  •  PDF-отчёт  •  Сценарный анализ  •  Docker",
        Inches(0.3), Inches(6.35), Inches(12.7), Inches(0.9),
        font_size=15, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)


def slide_architecture(prs):
    """Слайд 4 — Архитектура."""
    sl = blank_slide(prs)
    rect(sl, 0, 0, W, H, C_LIGHT)
    header_bar(sl, "Трёхуровневая архитектура",
               "ETL → ML/AI Core → Представление")

    # Уровни
    levels = [
        ("📂  Уровень данных",
         "data_loader.py  —  ETL Pipeline\nincidents.xlsx (220 строк, 50 колонок)\nkorgau_cards.xlsx (9 917 строк, 16 колонок)\nОбогащение: смена, стаж, флаги, resolved",
         C_DARK_BLUE),
        ("🧠  AI / ML Core + NLP",
         "Prophet  —  time-series forecast\nTF-IDF + KMeans  —  кластеры причин\nRisk Scoring  —  формула из ТЗ\nPearson / Spearman  —  корреляция\nGemini 2.0 Flash  —  NLP + Vision",
         C_BLUE),
        ("🖥️  Представление",
         "Streamlit  —  4 таба, 10 Plotly-графиков\nFastAPI  —  15 эндпоинтов, OpenAPI /docs\nfpdf2  —  PDF-отчёт 4 страницы\nAlert Manager  —  4 уровня CRITICAL→LOW",
         C_RED),
    ]

    for i, (title, body, color) in enumerate(levels):
        y = Inches(1.5 + i * 1.9)
        rect(sl, Inches(0.4), y, Inches(12.5), Inches(1.75), color)
        txt(sl, title, Inches(0.6), y + Inches(0.1), Inches(4.5), Inches(0.55),
            font_size=16, bold=True, color=C_WHITE)
        txt(sl, body, Inches(5.2), y + Inches(0.08), Inches(7.5), Inches(1.5),
            font_size=13, color=C_WHITE)

        # Разделитель
        if i < 2:
            txt(sl, "▼", Inches(6.3), y + Inches(1.78), Inches(0.8), Inches(0.3),
                font_size=14, color=C_DARK_BLUE, align=PP_ALIGN.CENTER)


def slide_predict(prs):
    """Слайд 5 — Предиктивная аналитика."""
    sl = blank_slide(prs)
    rect(sl, 0, 0, W, H, C_LIGHT)
    header_bar(sl, "Предиктивная аналитика",
               "Prophet Forecast  •  Risk Scoring  •  Сценарный анализ мер контроля")

    # Колонка 1 — Prophet
    rect(sl, Inches(0.3), Inches(1.4), Inches(4.1), Inches(5.8), C_WHITE)
    txt(sl, "🔮  Prophet Forecast", Inches(0.4), Inches(1.5), Inches(4.0), Inches(0.5),
        font_size=15, bold=True, color=C_DARK_BLUE)
    txb = sl.shapes.add_textbox(Inches(0.4), Inches(2.05), Inches(3.9), Inches(4.8))
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.add_run().text = ""
    for item in [
        "Прогноз на 12 / 6 / 3 месяца",
        "95% доверительный интервал",
        "Годовая сезонность",
        "Fallback: линейная регрессия",
        "Компонент тренда на графике",
    ]:
        add_bullet(tf, f"• {item}", size=13, color=C_DARK)

    # Колонка 2 — Risk Scoring
    rect(sl, Inches(4.65), Inches(1.4), Inches(4.1), Inches(5.8), C_WHITE)
    txt(sl, "⚡  Risk Scoring (ТЗ)", Inches(4.75), Inches(1.5), Inches(4.0), Inches(0.5),
        font_size=15, bold=True, color=C_DARK_BLUE)

    formula_lines = [
        "risk_score =",
        "  0.35 × incident_rate",
        "+ 0.30 × violation_rate",
        "+ 0.20 × trend_slope",
        "+ 0.15 × severity_factor",
        "",
        "Нормализация: 0–100",
        "CRITICAL ≥ 70",
        "HIGH     ≥ 50",
        "MEDIUM   ≥ 25",
        "LOW       < 25",
    ]
    txb2 = sl.shapes.add_textbox(Inches(4.75), Inches(2.05), Inches(3.9), Inches(4.8))
    tf2 = txb2.text_frame; tf2.word_wrap = True
    p2 = tf2.paragraphs[0]; p2.add_run().text = ""
    for line in formula_lines:
        add_bullet(tf2, line, size=12,
                   color=C_RED if line.startswith("CRITICAL") else
                         C_AMBER if line.startswith("HIGH") else
                         C_DARK)

    # Колонка 3 — Scenario
    rect(sl, Inches(9.0), Inches(1.4), Inches(4.1), Inches(5.8), C_WHITE)
    txt(sl, "🎯  Сценарный анализ", Inches(9.1), Inches(1.5), Inches(4.0), Inches(0.5),
        font_size=15, bold=True, color=C_DARK_BLUE)
    txb3 = sl.shapes.add_textbox(Inches(9.1), Inches(2.05), Inches(3.9), Inches(4.8))
    tf3 = txb3.text_frame; tf3.word_wrap = True
    p3 = tf3.paragraphs[0]; p3.add_run().text = ""
    for item in [
        "9 мер контроля (LOTO,",
        "  СИЗ, Высота, Огонь...)",
        "Снижение нарушений",
        "  по категориям Коргау",
        "× Pearson r = 0.415",
        "= ожид. снижение НС",
        "",
        "Экономика: избежанные",
        "инциденты × 5 млн ₸",
    ]:
        add_bullet(tf3, item, size=13, color=C_DARK)


def slide_korgau(prs):
    """Слайд 6 — Карты Коргау."""
    sl = blank_slide(prs)
    rect(sl, 0, 0, W, H, C_LIGHT)
    header_bar(sl, "Карты Коргау — поведенческий аудит безопасности",
               "9 917 наблюдений  •  Корреляция с инцидентами r = 0.415 (Pearson)")

    features = [
        ("F-10", "Паттерны нарушений", "groupby (org, category, month)\nтоп нарушителей и категорий"),
        ("F-11", "Система алертов", "CRITICAL > 50 нарушений/30 дней\nHIGH > 30  •  MEDIUM > 15  •  LOW > 5"),
        ("F-12", "Рейтинг организаций", "64 org  •  violation_rate %\ncompliance_rate  •  work_stops"),
        ("F-13", "Корреляция НС", "Pearson r=0.415  Spearman r=0.501\nстатистически значима"),
        ("F-14", "Pre-alert тренд", "рост / снижение / стабильно\nпо последним 3 месяцам"),
        ("F-15", "AI-кластеризация", "Gemini 2.0 Flash\n10 тематических кластеров"),
    ]

    for i, (fid, title, desc) in enumerate(features):
        col = i % 3
        row = i // 3
        x = Inches(0.3 + col * 4.35)
        y = Inches(1.5 + row * 2.55)
        rect(sl, x, y, Inches(4.1), Inches(2.3), C_DARK_BLUE if row == 0 else C_BLUE)
        txt(sl, fid, x + Inches(0.15), y + Inches(0.1), Inches(1.0), Inches(0.35),
            font_size=11, bold=True, color=C_AMBER)
        txt(sl, title, x + Inches(0.15), y + Inches(0.4), Inches(3.8), Inches(0.45),
            font_size=14, bold=True, color=C_WHITE)
        txt(sl, desc, x + Inches(0.15), y + Inches(0.85), Inches(3.8), Inches(1.3),
            font_size=12, color=RGBColor(0xba, 0xd4, 0xf8))


def slide_economic(prs):
    """Слайд 7 — Экономический эффект."""
    sl = blank_slide(prs)
    rect(sl, 0, 0, W, H, C_LIGHT)
    header_bar(sl, "Экономический эффект",
               "Измеримый результат внедрения AI-аналитики")

    # KPI в людях
    rect(sl, Inches(0.3), Inches(1.4), Inches(12.7), Inches(0.4), C_DARK_BLUE)
    txt(sl, "В ЛЮДЯХ", Inches(0.3), Inches(1.4), Inches(12.7), Inches(0.4),
        font_size=13, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    kpis = [
        ("↓ 38%",  "Несчастных\nслучаев"),
        ("↓ 40%",  "Микро-\nтравм"),
        ("↓ 50%",  "Опасных\nситуаций"),
        ("↓ 83%",  "Время\nреагирования"),
        ("↑ 70%",  "Охват\nпредиктивом"),
    ]
    for i, (val, lbl) in enumerate(kpis):
        x = Inches(0.3 + i * 2.56)
        color = C_GREEN if val.startswith("↑") else C_RED
        kpi_box(sl, val, lbl, x, Inches(1.85), Inches(2.45), Inches(1.35), bg=color)

    # В деньгах — таблица
    rect(sl, Inches(0.3), Inches(3.35), Inches(12.7), Inches(0.4), C_DARK_BLUE)
    txt(sl, "В ДЕНЬГАХ", Inches(0.3), Inches(3.35), Inches(12.7), Inches(0.4),
        font_size=13, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    rows = [
        ("Прямые затраты на НС  (7 × 5 млн ₸)",          "≈ 35 000 000 ₸"),
        ("Косвенные потери  (простой, репутация)",         "≈ 70 000 000 ₸"),
        ("Штрафы и предписания  (−30%)",                   "≈  5 000 000 ₸"),
        ("Расследования и отчётность  (−70%)",             "≈  8 000 000 ₸"),
        ("ROI автоматизации аудитов  (+40%)",              "≈  3 000 000 ₸"),
    ]
    for i, (label, value) in enumerate(rows):
        y = Inches(3.8 + i * 0.52)
        bg = C_WHITE if i % 2 == 0 else C_LIGHT
        rect(sl, Inches(0.3), y, Inches(12.7), Inches(0.5), bg)
        txt(sl, label, Inches(0.5), y + Inches(0.08), Inches(9.5), Inches(0.35),
            font_size=13, color=C_DARK)
        txt(sl, value, Inches(10.0), y + Inches(0.08), Inches(2.8), Inches(0.35),
            font_size=13, bold=True, color=C_DARK_BLUE, align=PP_ALIGN.RIGHT)

    # Итого
    rect(sl, Inches(0.3), Inches(6.45), Inches(12.7), Inches(0.7), C_GREEN)
    txt(sl, "ИТОГО:   ≈ 121 000 000 ₸  (~250 000 USD)  в год",
        Inches(0.3), Inches(6.45), Inches(12.7), Inches(0.7),
        font_size=20, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)


def slide_tech(prs):
    """Слайд 8 — Технологический стек."""
    sl = blank_slide(prs)
    rect(sl, 0, 0, W, H, C_LIGHT)
    header_bar(sl, "Технологический стек",
               "Production-ready Python-монолит с Docker Compose")

    stack = [
        ("🔧  Backend",    "Python 3.11  •  FastAPI  •  uvicorn\n15 REST-эндпоинтов  •  OpenAPI /docs",    C_DARK_BLUE),
        ("📊  ML Core",    "Prophet (time series)\nscikit-learn: TF-IDF + KMeans\nRisk Scoring + Pearson/Spearman",  C_BLUE),
        ("🤖  NLP / CV",   "Google Gemini 2.0 Flash\nклассификация  •  рекомендации\nGemini Vision (CV-анализ фото)",C_BLUE),
        ("📈  Frontend",   "Streamlit  •  4 таба\n10 Plotly-графиков\nCI-лента прогноза",                  C_DARK_BLUE),
        ("📄  Отчёты",     "fpdf2  •  4-страничный PDF\nDejaVu TTF (кириллица)\nKPI + графики + алерты",   C_RED),
        ("🐳  DevOps",     "Docker + Docker Compose\nhse-api :8002  •  hse-ui :8502\nhealthcheck + networks",C_RED),
    ]

    for i, (title, body, color) in enumerate(stack):
        col = i % 3
        row = i // 2
        x = Inches(0.3 + col * 4.35)
        y = Inches(1.5 + row * 2.7)
        rect(sl, x, y, Inches(4.1), Inches(2.45), color)
        txt(sl, title, x + Inches(0.15), y + Inches(0.12), Inches(3.8), Inches(0.45),
            font_size=15, bold=True, color=C_WHITE)
        rect(sl, x, y + Inches(0.58), Inches(4.1), Inches(0.03), C_WHITE)
        txt(sl, body, x + Inches(0.15), y + Inches(0.68), Inches(3.8), Inches(1.6),
            font_size=13, color=RGBColor(0xba, 0xd4, 0xf8))


def slide_demo(prs):
    """Слайд 9 — Дашборд / скриншоты."""
    sl = blank_slide(prs)
    rect(sl, 0, 0, W, H, C_LIGHT)
    header_bar(sl, "Дашборд — 4 вкладки Streamlit",
               "http://localhost:8502  •  Swagger http://localhost:8002/docs")

    tabs = [
        ("📊 Инциденты",
         ["Динамика + скользящее среднее 3м",
          "Pie по типам инцидентов",
          "Топ-10 организаций",
          "Кластеры причин (TF-IDF + KMeans)",
          "NLP-классификатор (Gemini)",
          "CV-анализ фото (Gemini Vision)"],
         C_DARK_BLUE),
        ("🔍 Карты Коргау",
         ["Тренд нарушений топ-5 org",
          "Рейтинг организаций (% нарушений)",
          "Хорошие vs плохие практики",
          "Pearson r / Spearman r метрики"],
         C_BLUE),
        ("🔮 Предикт",
         ["Prophet: прогноз + 95% CI лента",
          "Риск-скоринг + heatmap (0–100)",
          "Матрица корреляций (heatmap)",
          "Сценарный анализ мер контроля"],
         C_BLUE),
        ("🚨 Алерты",
         ["CRITICAL/HIGH/MEDIUM/LOW карточки",
          "Фильтры по org и уровню",
          "Счётчики по уровням",
          "Кнопка скачать PDF-отчёт"],
         C_RED),
    ]

    for i, (title, items, color) in enumerate(tabs):
        x = Inches(0.3 + i * 3.2)
        rect(sl, x, Inches(1.4), Inches(3.0), Inches(5.8), color)
        txt(sl, title, x + Inches(0.1), Inches(1.5), Inches(2.8), Inches(0.5),
            font_size=14, bold=True, color=C_WHITE)
        rect(sl, x, Inches(2.0), Inches(3.0), Inches(0.03), C_WHITE)
        txb = sl.shapes.add_textbox(x + Inches(0.1), Inches(2.1), Inches(2.8), Inches(4.8))
        tf = txb.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.add_run().text = ""
        for item in items:
            add_bullet(tf, f"✓  {item}", size=12, color=C_WHITE)


def slide_results(prs):
    """Слайд 10 — Итог / CTA."""
    sl = blank_slide(prs)
    rect(sl, 0, 0, W, H, C_DARK_BLUE)
    rect(sl, 0, Inches(5.6), W, Inches(1.9), C_DARK)
    rect(sl, 0, Inches(5.5), W, Inches(0.06), C_RED)

    txt(sl, "🛡️", Inches(0.5), Inches(0.5), Inches(1.2), Inches(1.2),
        font_size=56, color=C_WHITE)

    txt(sl, "Результат",
        Inches(1.8), Inches(0.5), Inches(10), Inches(0.9),
        font_size=46, bold=True, color=C_WHITE)

    checklist = [
        "✅  Рабочий AI-дашборд  —  аналитика прошлых инцидентов",
        "✅  Предикт на 12 месяцев  —  Prophet + 95% CI",
        "✅  Сценарный анализ  —  расчёт эффекта мер контроля",
        "✅  4-уровневые алерты  —  CRITICAL → LOW",
        "✅  Топ-5 AI-рекомендаций  —  Gemini 2.0 Flash",
        "✅  CV-анализ фото  —  Gemini Vision",
        "✅  PDF-отчёт  —  4 страницы, кириллица",
        "✅  Экономика  —  121 млн ₸ / год",
    ]
    txb = sl.shapes.add_textbox(Inches(1.8), Inches(1.5), Inches(11), Inches(3.8))
    tf = txb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.add_run().text = ""
    for item in checklist:
        add_bullet(tf, item, size=15,
                   color=C_WHITE if "✅" in item else RGBColor(0xba, 0xd4, 0xf8))

    txt(sl, "FastAPI :8002  •  Streamlit :8502  •  docker compose up --build",
        Inches(0.5), Inches(5.8), Inches(12), Inches(0.5),
        font_size=14, color=RGBColor(0xba, 0xd4, 0xf8), align=PP_ALIGN.CENTER)

    txt(sl, "KMG Hackathon 2026  •  github.com/Islambek201632838/kmg-hse",
        Inches(0.5), Inches(6.5), Inches(12), Inches(0.5),
        font_size=13, color=RGBColor(0x64, 0x74, 0x8b), align=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def build():
    prs = new_prs()

    slide_title(prs)
    slide_problem(prs)
    slide_solution(prs)
    slide_architecture(prs)
    slide_predict(prs)
    slide_korgau(prs)
    slide_economic(prs)
    slide_tech(prs)
    slide_demo(prs)
    slide_results(prs)

    out = "HSE_Presentation.pptx"
    prs.save(out)
    print(f"✅  Презентация сохранена: {out}  ({len(prs.slides)} слайдов)")


if __name__ == "__main__":
    build()