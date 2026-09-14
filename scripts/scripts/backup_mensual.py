"""
FAMM - Respaldo mensual con fecha
-----------------------------------
Corre el dia 15 de cada mes. Descarga el Excel original completo
(tal cual, sin filtrar columnas -- es un respaldo, no el dashboard
publico) y lo sube a una carpeta de "Archivo historico" en Drive
con el nombre AAAAMMDD_gastos.xlsx

Requiere:
  - GOOGLE_SERVICE_ACCOUNT_JSON (con permiso de lectura sobre el
    archivo original y de escritura sobre la carpeta de respaldo)
  - SPREADSHEET_ID
  - BACKUP_FOLDER_ID (carpeta de Drive donde viven los respaldos)
"""

import os
import io
import json
import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]

SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
BACKUP_FOLDER_ID = os.environ["BACKUP_FOLDER_ID"]

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_drive_service():
    creds_json = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def main():
    service = get_drive_service()

    # Exporta el Google Sheet original como .xlsx
    contenido = service.files().export(
        fileId=SPREADSHEET_ID, mimeType=XLSX_MIME
    ).execute()

    hoy = datetime.date.today()
    nombre_archivo = f"{hoy.strftime('%Y%m%d')}_gastos.xlsx"

    media = MediaIoBaseUpload(
        io.BytesIO(contenido), mimetype=XLSX_MIME, resumable=True
    )
    metadata = {"name": nombre_archivo, "parents": [BACKUP_FOLDER_ID]}

    archivo = service.files().create(
        body=metadata, media_body=media, fields="id, name"
    ).execute()

    print(f"Respaldo creado: {archivo['name']} (id: {archivo['id']})")


if __name__ == "__main__":
    main()
