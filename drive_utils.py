"""Utilidades para conectar Streamlit Cloud con Google Drive mediante OAuth2."""
import io
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import streamlit as st


def _get_service():
  oauth = st.secrets["google_oauth"]
  creds = Credentials(
      token=None,
      refresh_token=oauth["refresh_token"],
      token_uri="https://oauth2.googleapis.com/token",
      client_id=oauth["client_id"],
      client_secret=oauth["client_secret"],
      scopes=["https://www.googleapis.com/auth/drive"],
  )
  creds.refresh(Request())
  return build("drive", "v3", credentials=creds, cache_discovery=False)


def descargar(file_id: str) -> bytes:
  service = _get_service()
  request = service.files().get_media(fileId=file_id)
  buf = io.BytesIO()
  downloader = MediaIoBaseDownload(buf, request)
  done = False
  while not done:
    _, done = downloader.next_chunk()
  return buf.getvalue()


def actualizar(file_id: str, contenido: bytes, mime_type: str):
  service = _get_service()
  media = MediaIoBaseUpload(
      io.BytesIO(contenido), mimetype=mime_type, resumable=True
  )
  service.files().update(fileId=file_id, media_body=media).execute()


def copiar(file_id: str, nuevo_nombre: str, folder_id: str):
  service = _get_service()
  body = {"name": nuevo_nombre, "parents": [folder_id]}
  service.files().copy(fileId=file_id, body=body).execute()


def subir_nuevo(
    nombre: str,
    contenido: bytes,
    folder_id: str,
    mime_type: str = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
):
  service = _get_service()
  body = {"name": nombre, "parents": [folder_id]}
  media = MediaIoBaseUpload(
      io.BytesIO(contenido), mimetype=mime_type, resumable=True
  )
  service.files().create(body=body, media_body=media).execute()
