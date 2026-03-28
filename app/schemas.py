from pydantic import BaseModel
from typing import Any, Optional


class ClassifyRequest(BaseModel):
    description: str


class ClassifyResponse(BaseModel):
    type: str
    severity: str
    root_cause_category: str
    summary: str


class AlertItem(BaseModel):
    org: str
    level: str
    color: str
    count_30d: int
    top_categories: dict[str, int]
    period: str
    repeated_violation_reason: Optional[str] = None
    yoy_reason: Optional[str] = None


class RiskScoreItem(BaseModel):
    org: str
    risk_score: float
    risk_level: str
    incident_count: float
    violation_count: float


class ForecastPoint(BaseModel):
    ds: Any
    yhat: float
    yhat_lower: float
    yhat_upper: float


class RecommendationItem(BaseModel):
    priority: str
    action: str
    rationale: str
    expected_effect: str


class ScenarioRequest(BaseModel):
    org: str
    measures: list[str]
