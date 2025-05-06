from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import StreamingResponse
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import io, pathlib

from generation import run_reporter

app = FastAPI(title="Form-Reporter API")

@app.post("/generate")
async def generate(
    excel: UploadFile = File(...),
    template: UploadFile = File(...)
):
    # vérification du type
    if not excel.filename.lower().endswith((".xls", ".xlsx")):
        raise HTTPException(400, "Format Excel invalide")
    if not template.filename.lower().endswith(".docx"):
        raise HTTPException(400, "Format Word invalide")

    with TemporaryDirectory() as tmp:
        tmpdir = pathlib.Path(tmp)
        excel_fp  = tmpdir / excel.filename
        tpl_fp    = tmpdir / template.filename
        # écriture sur disque
        with open(excel_fp, "wb") as f:
            f.write(await excel.read())
        with open(tpl_fp, "wb") as f:
            f.write(await template.read())

        out_dir = tmpdir / "output"
        run_reporter(excel_fp, tpl_fp, out_dir)

        # empaqueter en ZIP
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
