from fastapi import APIRouter, UploadFile, File, HTTPException
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
