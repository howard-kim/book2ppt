"""
FastAPI backend
================
POST /convert  — accepts .idml upload, returns .pptx download

Deploy on Railway / Render / Fly.io.
Place your template.pptx in the same directory as this file.
"""

from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response

from idml_parser import parse_idml
from pptx_generator import generate_ppt

import tempfile

# ---------------------------------------------------------------------------
app = FastAPI(title="IDML → PPTX Converter")

# CORS: allow Vercel frontend (configure ALLOWED_ORIGINS in production)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],   # tighten to your Vercel URL before going live
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
            raise HTTPException(status_code=422, detail=str(e))

        pptx_bytes = output_path.read_bytes()

    stem = Path(file.filename).stem
    return Response(
        content=pptx_bytes,
        media_type=(
            "application/vnd.openxmlformats-officedocument"
            ".presentationml.presentation"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{stem}.pptx"'
        },
    )
