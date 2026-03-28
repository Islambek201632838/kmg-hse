from fastapi import APIRouter
from fastapi.responses import StreamingResponse
import io
from app.services.report_generator import generate_pdf, generate_excel

router = APIRouter(prefix="/api/reports", tags=["reports"])


@router.get("/pdf", summary="Скачать PDF-отчёт")
def download_pdf():
    pdf_bytes = generate_pdf()
    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=hse_report.pdf"},
    )


@router.get("/excel", summary="Скачать Excel-отчёт (F-09)")
def download_excel():
    excel_bytes = generate_excel()
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=hse_report.xlsx"},
    )
