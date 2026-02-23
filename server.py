from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import shutil
import os
import uuid
import subprocess
import zipfile

app = FastAPI()

@app.get("/")
def home():
    return {"status": "online"}

@app.post("/analisar")
async def analisar(file: UploadFile = File(...)):

    temp_id = str(uuid.uuid4())
    base_path = f"/tmp/{temp_id}"
    os.makedirs(base_path, exist_ok=True)

    zip_path = os.path.join(base_path, "cliente.zip")

    # Salva zip enviado
    with open(zip_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    # Descompacta
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(base_path)

    # Executa o bot
    subprocess.run(["python", "bot.py", base_path])

    resultado = os.path.join(base_path, "ANALISE FINAL.xlsx")

    if os.path.exists(resultado):
        return FileResponse(resultado, filename="ANALISE FINAL.xlsx")

    return {"erro": "Falha na análise"}