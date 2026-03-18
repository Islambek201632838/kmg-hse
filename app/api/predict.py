from fastapi import APIRouter
from app.schemas import RiskScoreItem, ScenarioRequest
from app.services.predictor import get_forecast, get_risk_scores, get_correlation_matrix, calculate_scenario, CONTROL_MEASURES

router = APIRouter(prefix="/api/predict", tags=["predict"])


@router.get("/forecast")
def forecast(periods: int = 12):
    return get_forecast(periods=periods)


@router.get("/risk-scores")
def risk_scores():
    return get_risk_scores()


@router.get("/correlation-matrix")
def correlation_matrix():
    return get_correlation_matrix()


@router.get("/scenario/measures")
def available_measures():
    """Список доступных мер контроля для Scenario Modeling."""
    return {"measures": list(CONTROL_MEASURES.keys())}


@router.post("/scenario")
def scenario(req: ScenarioRequest):
    """
    Расчёт изменения вероятности инцидентов при внедрении мер контроля.
    Возвращает: baseline, predicted, reduction %, avoided incidents, экономический эффект.
    """
    return calculate_scenario(org=req.org, measures=req.measures)
