from fastapi import APIRouter
from fastapi.responses import StreamingResponse
import io
from app.services.report_generator import generate_pdf

router = APIRouter(prefix="/api/reports", tags=["reports"])


@router.get("/pdf")
def download_pdf():
    pdf_bytes = generate_pdf()
    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=hse_report.pdf"},
    )
