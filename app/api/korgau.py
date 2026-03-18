from fastapi import APIRouter
from app.schemas import AlertItem, RiskScoreItem
from app.services.korgau_analyzer import (
    get_alerts, get_org_rankings, get_violation_rates,
    get_org_trends, get_good_vs_bad, get_correlation_with_incidents,
)

router = APIRouter(prefix="/api/korgau", tags=["korgau"])


@router.get("/alerts", response_model=list[AlertItem])
def korgau_alerts():
    return get_alerts()


@router.get("/rankings")
def korgau_rankings():
    return get_org_rankings()


@router.get("/violation-rates")
def violation_rates():
    return get_violation_rates()


@router.get("/trends")
def org_trends():
    return get_org_trends()


@router.get("/good-vs-bad")
def good_vs_bad():
    return get_good_vs_bad()


@router.get("/correlation")
def correlation():
    return get_correlation_with_incidents()
