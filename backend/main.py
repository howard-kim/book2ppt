"""
FastAPI backend
================
POST /convert  — accepts .idml upload, returns .pptx download

Deploy on Railway / Render / Fly.io.
Place your template.pptx in the same directory as this file.
"""

import os
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response

from idml_parser import parse_idml
from pptx_generator import generate_ppt

import tempfile

# ---------------------------------------------------------------------------
app = FastAPI(title="IDML → PPTX Converter")


def get_allowed_origins() -> list[str]:
    """
    Read CORS origins from ALLOWED_ORIGINS.

    Example:
    ALLOWED_ORIGINS=https://my-frontend.app,https://www.my-frontend.app
    """
    raw = os.getenv("ALLOWED_ORIGINS", "*").strip()
    if not raw:
        return ["*"]
    if raw == "*":
        return ["*"]
    return [origin.strip() for origin in raw.split(",") if origin.strip()]

# CORS: allow Vercel frontend (configure ALLOWED_ORIGINS in production)
app.add_middleware(
    CORSMiddleware,
    allow_origins=get_allowed_origins(),
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

TEMPLATE_PATH = Path(__file__).parent / "template.pptx"


# ---------------------------------------------------------------------------
@app.get("/health")
def health():
    """Quick liveness check."""
    return {"status": "ok"}


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    """
    Accept an .idml file and return a .pptx file.

    Request : multipart/form-data  field name = "file"
    Response: application/vnd.openxmlformats-officedocument.presentationml.presentation
    """
    if not file.filename.lower().endswith(".idml"):
        raise HTTPException(status_code=400, detail="Only .idml files are accepted.")

    if not TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail="template.pptx not found on server. Contact the administrator.",
        )

    content = await file.read()

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        idml_path   = tmp / "input.idml"
        output_path = tmp / "output.pptx"

        idml_path.write_bytes(content)

        try:
            data = parse_idml(str(idml_path))
            generate_ppt(data, str(TEMPLATE_PATH), str(output_path))
        except Exception as e:
            import traceback; traceback.print_exc()
            raise HTTPException(status_code=422, detail=str(e))

        pptx_bytes = output_path.read_bytes()

    stem = Path(file.filename).stem
    from urllib.parse import quote
    encoded_name = quote(f"{stem}.pptx", safe="")
    return Response(
        content=pptx_bytes,
        media_type=(
            "application/vnd.openxmlformats-officedocument"
            ".presentationml.presentation"
        ),
        headers={
            "Content-Disposition": f"attachment; filename=\"output.pptx\"; filename*=UTF-8''{encoded_name}"
        },
    )
