import os
import json
import tempfile
from pathlib import Path

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import FileResponse

from generar_contrato import generar_contrato  # tu script

# Carpeta base del proyecto
BASE_DIR = Path(__file__).resolve().parent

# Ruta fija a la plantilla y carpeta de salida
PLANTILLA_PATH = BASE_DIR / "plantilla.docx"
SALIDA_DIR = BASE_DIR / "salida_contratos"
SALIDA_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Servicio de contratos")

@app.post("/contrato")
async def crear_contrato(request: Request):
    """
    Recibe un JSON con los datos del contrato y genera un .docx
    usando tu función generar_contrato.
    Devuelve una URL para descargar el archivo.
    """
    try:
        datos = await request.json()
    except Exception:
        raise HTTPException(status_code=400, detail="Body debe ser JSON válido")

    # Guardar JSON en un archivo temporal porque generar_contrato
    # espera una ruta a archivo de datos
    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".json",
        delete=False,
        dir=SALIDA_DIR
    ) as tmp:
        json.dump(datos, tmp, ensure_ascii=False, indent=2)
        tmp_path = tmp.name

    # Llamar a tu función
    out_path = generar_contrato(
        datos_path=tmp_path,
        plantilla_path=str(PLANTILLA_PATH),
        salida_dir=str(SALIDA_DIR),
        listar_marcadores=False,
        resaltar=False,
    )

    # Si por alguna razón no devolvió ruta
    if not out_path:
        raise HTTPException(status_code=500, detail="No se generó el contrato")

    filename = os.path.basename(out_path)

    # Construir URL pública al archivo usando la propia app
    file_url = str(request.url_for("descargar_contrato", filename=filename))

    return {"fileUrl": file_url}

@app.get("/files/{filename}", name="descargar_contrato")
async def descargar_contrato(filename: str):
    """
    Devuelve el archivo .docx generado.
    """
    file_path = SALIDA_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Archivo no encontrado")

    return FileResponse(
        path=str(file_path),
        media_type=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        filename=filename,
    )
