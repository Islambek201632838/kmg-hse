from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.api import incidents, korgau, predict, reports

app = FastAPI(
    title="HSE AI Analytics",
    description="AI-аналитика охраны труда: инциденты, Карты Коргау, предиктив",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(incidents.router)
app.include_router(korgau.router)
app.include_router(predict.router)
app.include_router(reports.router)


@app.get("/health")
def health():
    return {"status": "ok"}
