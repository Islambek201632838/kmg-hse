import json
import re
import functools
from google import genai
from app.config import settings

_client = genai.Client(api_key=settings.gemini_api_key)


def _parse_json(text: str) -> dict | list:
    """Надёжный парсинг JSON из ответа Gemini (убирает markdown-блоки)."""
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return json.loads(text.strip())


def _generate(prompt: str) -> str:
    response = _client.models.generate_content(
        model=settings.gemini_model,
        contents=prompt,
    )
    return response.text


@functools.lru_cache(maxsize=512)
def classify_incident(description: str) -> str:
    """Классифицировать инцидент через Gemini. Кэшируется по тексту описания."""
    prompt = f"""Ты — эксперт по охране труда в нефтегазовой отрасли.
Классифицируй инцидент по описанию. Верни ТОЛЬКО JSON без markdown:
{{
  "type": "<НС|Микротравма|Опасная ситуация|ДТП|Пожар|Экологический инцидент|Прочее>",
  "severity": "<Критический|Высокий|Средний|Низкий>",
  "root_cause_category": "<Человеческий фактор|Оборудование|Процедуры|Среда|Организация>",
  "summary": "<краткое резюме на русском, 1-2 предложения>"
}}

Описание: {description}"""
    return _generate(prompt)


def classify_incident_parsed(description: str) -> dict:
    try:
        return _parse_json(classify_incident(description))
    except Exception:
        return {
            "type": "Прочее",
            "severity": "Средний",
            "root_cause_category": "Неизвестно",
            "summary": description[:120],
        }


def classify_incidents_batch(descriptions: list[str]) -> list[dict]:
    """Батчевая классификация. LRU-кэш не повторяет одинаковые описания."""
    return [classify_incident_parsed(str(d)) for d in descriptions]


def generate_recommendations(patterns: dict) -> list[dict]:
    """Топ-5 рекомендаций по мерам контроля на основе паттернов инцидентов."""
    prompt = f"""Ты — эксперт HSE в нефтегазовой отрасли.
На основе данных сформируй топ-5 конкретных рекомендаций по снижению травматизма.
Верни ТОЛЬКО JSON без markdown — список:
[
  {{
    "priority": "<Высокий|Средний|Низкий>",
    "action": "<конкретное действие>",
    "rationale": "<обоснование на основе данных>",
    "expected_effect": "<ожидаемый эффект>"
  }}
]

Данные:
- Топ организации по инцидентам: {patterns.get("top_orgs", [])}
- Топ типы инцидентов: {patterns.get("top_types", [])}
- Топ причины: {patterns.get("top_causes", [])}
- Тренд (рост/снижение): {patterns.get("trend", "нет данных")}"""

    try:
        return _parse_json(_generate(prompt))
    except Exception:
        return []


@functools.lru_cache(maxsize=2048)
def classify_korgau_observation(description: str, obs_type: str) -> str:
    """Классифицировать наблюдение Карты Коргау. Кэшируется."""
    prompt = f"""Ты — эксперт по охране труда.
Классифицируй наблюдение поведенческого аудита безопасности. Верни ТОЛЬКО JSON без markdown:
{{
  "cluster": "<СИЗ|Работа на высоте|LOTO|Пожарная безопасность|Транспорт|Электробезопасность|Химическая безопасность|Организация рабочего места|Процедуры и разрешения|Другое>",
  "risk_level": "<Критический|Высокий|Средний|Низкий>",
  "tags": ["<тег1>", "<тег2>"]
}}

Тип: {obs_type}
Описание: {description[:600]}"""
    return _generate(prompt)


def classify_korgau_parsed(description: str, obs_type: str) -> dict:
    try:
        return _parse_json(classify_korgau_observation(description, obs_type))
    except Exception:
        return {"cluster": "Другое", "risk_level": "Средний", "tags": []}


def analyze_incident_photo(image_bytes: bytes, mime_type: str = "image/jpeg") -> dict:
    """F-08: CV-анализ фото с места происшествия через Gemini Vision."""
    from google.genai import types as genai_types

    prompt = """Ты — эксперт по охране труда и промышленной безопасности.
Проанализируй фотографию с места происшествия или производственного объекта.
Верни ТОЛЬКО JSON без markdown:
{
  "risk_level": "<Критический|Высокий|Средний|Низкий>",
  "violations": [
    {"type": "<тип нарушения>", "description": "<детальное описание>", "location": "<где на фото>"}
  ],
  "missing_ppe": ["<СИЗ которых не хватает>"],
  "unsafe_conditions": ["<опасные условия>"],
  "unsafe_actions": ["<опасные действия персонала>"],
  "recommendations": ["<конкретная рекомендация>"],
  "summary": "<общее резюме анализа на русском, 2-3 предложения>"
}
Если фото не связано с производством или охраной труда — верни risk_level: Низкий и пустые списки."""

    try:
        response = _client.models.generate_content(
            model=settings.gemini_model,
            contents=[
                genai_types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
                prompt,
            ],
        )
        return _parse_json(response.text)
    except Exception as e:
        return {
            "risk_level": "Неизвестно",
            "violations": [],
            "missing_ppe": [],
            "unsafe_conditions": [],
            "unsafe_actions": [],
            "recommendations": [],
            "summary": f"Ошибка анализа: {str(e)[:100]}",
        }
