from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import io, pathlib

from generation import run_reporter

app = FastAPI(title="Form-Reporter API")

# --- CORS pour autoriser ton front Netlify ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://silly-chaja-152562.netlify.app"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
# ----------------------------------------------

@app.post("/generate")
async def generate(
    excel: UploadFile = File(...),
    template: UploadFile = File(...)
):
    if not excel.filename.lower().endswith((".xls", ".xlsx")):
        raise HTTPException(400, "Format Excel invalide")
    if not template.filename.lower().endswith(".docx"):
        raise HTTPException(400, "Format Word invalide")

    with TemporaryDirectory() as tmp:
        tmpdir = pathlib.Path(tmp)
        excel_fp = tmpdir / excel.filename
        tpl_fp   = tmpdir / template.filename

        # Sauvegarde les uploads
        with open(excel_fp, "wb") as f:
            f.write(await excel.read())
        with open(tpl_fp, "wb") as f:
            f.write(await template.read())

        out_dir = tmpdir / "output"
        run_reporter(excel_fp, tpl_fp, out_dir)

        # Crée le ZIP en mémoire
        buf = io.BytesIO()
        with ZipFile(buf, "w") as zipf:
            for file in out_dir.glob("*.docx"):
                zipf.write(file, file.name)
        buf.seek(0)

        return StreamingResponse(
            buf,
            media_type="application/zip",
            headers={"Content-Disposition": "attachment; filename=reports.zip"}
        )
