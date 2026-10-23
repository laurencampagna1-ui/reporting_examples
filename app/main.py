from __future__ import annotations

import io
import re
import tempfile
import zipfile
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from . import data as data_mod
from . import pdf_convert, pptx_builder

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_TEMPLATE_PATH = BASE_DIR / "templates" / "base_template.pptx"

app = FastAPI(title="Reporting Demo")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "app" / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "app" / "templates"))


def _safe_filename(name: str) -> str:
    name = re.sub(r"[^\w\-. ]", "", name).strip()
    return name or "team"


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html")


@app.post("/generate")
async def generate(
    spreadsheet: UploadFile = File(...),
    report_period: str = Form(default="This Quarter"),
    output_formats: list[str] = Form(default=["pptx", "pdf"]),
):
    sheet_bytes = await spreadsheet.read()
    try:
        df = data_mod.load_spreadsheet(spreadsheet.filename or "upload.csv", sheet_bytes)
        teams, warnings = data_mod.build_team_decks(df)
    except data_mod.SpreadsheetError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    want_pptx = "pptx" in output_formats
    want_pdf = "pdf" in output_formats
    if not want_pptx and not want_pdf:
        raise HTTPException(status_code=400, detail="Select at least one output format.")

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for team_name, deck in teams.items():
            fname_base = _safe_filename(team_name)
            try:
                pptx_buf = pptx_builder.build_team_deck(
                    str(DEFAULT_TEMPLATE_PATH),
                    team_name,
                    report_period,
                    deck.in_progress_rows,
                    deck.completed_rows,
                )
            except Exception as exc:
                raise HTTPException(
                    status_code=500,
                    detail=f"Failed generating deck for '{team_name}': {exc}",
                ) from exc

            pptx_bytes = pptx_buf.getvalue()
            if want_pptx:
                zf.writestr(f"{fname_base}/{fname_base}.pptx", pptx_bytes)
            if want_pdf:
                try:
                    pdf_bytes = pdf_convert.pptx_bytes_to_pdf_bytes(pptx_bytes)
                except pdf_convert.PdfConversionError as exc:
                    raise HTTPException(status_code=500, detail=str(exc)) from exc
                zf.writestr(f"{fname_base}/{fname_base}.pdf", pdf_bytes)

        if warnings:
            zf.writestr("warnings.txt", "\n".join(warnings))

    zip_buf.seek(0)

    return StreamingResponse(
        zip_buf,
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=reports.zip"},
    )


@app.get("/health")
async def health():
    return {"status": "ok"}
