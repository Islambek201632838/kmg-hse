from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from typing import Optional
from app.schemas import ClassifyRequest, ClassifyResponse, RecommendationItem
from app.services.incident_analyzer import get_summary, get_cause_clusters, get_patterns, get_org_breakdown
from app.services.nlp_service import classify_incident_parsed, generate_recommendations, analyze_incident_photo

router = APIRouter(prefix="/api/incidents", tags=["incidents"])


@router.get("/summary")
def incidents_summary():
    return get_summary()


@router.post("/classify", response_model=ClassifyResponse)
def classify_incident(req: ClassifyRequest):
    return classify_incident_parsed(req.description)


@router.get("/cause-clusters")
def cause_clusters(n: int = 6):
    return get_cause_clusters(n_clusters=n)


@router.get("/org-breakdown")
def org_breakdown():
    return get_org_breakdown()


@router.get("/recommendations", response_model=list[RecommendationItem])
def recommendations():
    patterns = get_patterns()
    return generate_recommendations(patterns)


@router.post("/analyze-photo")
async def analyze_photo(file: UploadFile = File(...)):
    """F-08: CV-анализ фото с места происшествия (Gemini Vision)."""
    if not file.content_type.startswith("image/"):
        raise HTTPException(status_code=400, detail="Файл должен быть изображением")
    image_bytes = await file.read()
    return analyze_incident_photo(image_bytes, mime_type=file.content_type)


@router.post("/analyze", summary="Комбинированный анализ: текст и/или фото (F-01 + F-08)")
async def analyze_incident(
    description: Optional[str] = Form(None),
    file: Optional[UploadFile] = File(None),
):
    """
    Принимает текст и/или фото. Возвращает объединённый результат:
    - NLP-классификация текста (тип, тяжесть, причина)
    - CV-анализ фото (нарушения, СИЗ, рекомендации)
    Хотя бы одно из двух обязательно.
    """
    if not description and not file:
        raise HTTPException(400, "Укажите описание инцидента и/или загрузите фото")

    result = {}

    # NLP-классификация текста
    if description and description.strip():
        result["nlp"] = classify_incident_parsed(description.strip())

    # CV-анализ фото
    if file:
        if not file.content_type.startswith("image/"):
            raise HTTPException(400, "Файл должен быть изображением (JPG/PNG)")
        image_bytes = await file.read()
        result["cv"] = analyze_incident_photo(image_bytes, mime_type=file.content_type)

    return result
