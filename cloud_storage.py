from __future__ import annotations

import hashlib
import io
import os
from typing import Any

import requests


def setting(name: str) -> str:
    value = os.getenv(name, "")
    if value:
        return value
    try:
        import streamlit as st
        return str(st.secrets.get(name, ""))
    except Exception:
        return ""


def configured() -> bool:
    return bool(setting("SUPABASE_URL") and setting("SUPABASE_SERVICE_ROLE_KEY") and setting("GOOGLE_CLIENT_ID") and setting("GOOGLE_CLIENT_SECRET") and setting("GOOGLE_REFRESH_TOKEN"))


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def drive_service():
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build
    credentials = Credentials(
        token=None,
        refresh_token=setting("GOOGLE_REFRESH_TOKEN"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=setting("GOOGLE_CLIENT_ID"),
        client_secret=setting("GOOGLE_CLIENT_SECRET"),
        scopes=["https://www.googleapis.com/auth/drive.file"],
    )
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def ensure_folder(service, name: str, parent_id: str | None = None) -> str:
    query = "name = '{}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false".format(name.replace("'", "\\'"))
    if parent_id:
        query += f" and '{parent_id}' in parents"
    found = service.files().list(q=query, spaces="drive", fields="files(id,name)").execute().get("files", [])
    if found:
        return found[0]["id"]
    body: dict[str, Any] = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    if parent_id:
        body["parents"] = [parent_id]
    return service.files().create(body=body, fields="id").execute()["id"]


def upload_bytes(service, name: str, data: bytes, mime_type: str, folder_id: str) -> tuple[str, str]:
    from googleapiclient.http import MediaIoBaseUpload
    metadata = {"name": name, "parents": [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime_type, resumable=False)
    file = service.files().create(body=metadata, media_body=media, fields="id,webViewLink").execute()
    return file["id"], file.get("webViewLink", f"https://drive.google.com/open?id={file['id']}")


def update_bytes(service, file_id: str, data: bytes, mime_type: str) -> tuple[str, str]:
    from googleapiclient.http import MediaIoBaseUpload
    media = MediaIoBaseUpload(io.BytesIO(data), mimetype=mime_type, resumable=False)
    file = service.files().update(fileId=file_id, media_body=media, fields="id,webViewLink").execute()
    return file["id"], file.get("webViewLink", f"https://drive.google.com/open?id={file['id']}")


def supabase_request(path: str, method: str = "GET", payload: Any = None, params: dict[str, str] | None = None) -> Any:
    base = setting("SUPABASE_URL").rstrip("/")
    headers = {
        "apikey": setting("SUPABASE_SERVICE_ROLE_KEY"),
        "Authorization": f"Bearer {setting('SUPABASE_SERVICE_ROLE_KEY')}",
        "Content-Type": "application/json",
        "Prefer": "return=representation,resolution=merge-duplicates",
    }
    response = requests.request(method, f"{base}/rest/v1/{path}", headers=headers, params=params, json=payload, timeout=30)
    response.raise_for_status()
    return response.json() if response.content else None


def persist_analysis(*, client_name: str, farm_name: str, year: int, radius_m: float, original: bytes, shp_bytes: bytes, reanalysis: bytes, original_name: str, area_ha: float, point_count: int, outlier_count: int, attention_count: int) -> dict[str, Any] | None:
    if not configured():
        return None
    source_hash = sha256(original + shp_bytes)
    service = drive_service()
    existing = supabase_request("analyses", params={"source_hash": f"eq.{source_hash}", "select": "id,drive_folder_id"})
    existing_analysis = existing[0] if existing else None
    existing_files = {}
    if existing_analysis:
        file_rows = supabase_request("analysis_files", params={"analysis_id": f"eq.{existing_analysis['id']}", "select": "file_kind,drive_file_id"}) or []
        existing_files = {row["file_kind"]: row["drive_file_id"] for row in file_rows}
    root = (existing_analysis or {}).get("drive_folder_id") or ensure_folder(service, "Outlieria")
    originals = ensure_folder(service, "01_Fontes_Originais", root)
    reanalyses = ensure_folder(service, "02_Reanalises_Individuais", root)
    excel_id, excel_url = (update_bytes(service, existing_files["excel_original"], original, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") if "excel_original" in existing_files else upload_bytes(service, original_name, original, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", originals))
    shp_id, shp_url = (update_bytes(service, existing_files["shp"], shp_bytes, "application/octet-stream") if "shp" in existing_files else upload_bytes(service, f"{farm_name}_limite.shp", shp_bytes, "application/octet-stream", originals))
    reanalysis_id, reanalysis_url = (update_bytes(service, existing_files["reanálise_2026"], reanalysis, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") if "reanálise_2026" in existing_files else upload_bytes(service, f"{farm_name}_reanálise_{year}.xlsx", reanalysis, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", reanalyses))
    analysis = supabase_request("analyses?on_conflict=source_hash", "POST", {
        "client_name": client_name, "farm_name": farm_name, "analysis_year": year, "radius_m": round(radius_m, 2),
        "area_2026_ha": round(area_ha, 2), "point_count": point_count,
        "outlier_count": outlier_count, "attention_count": attention_count,
        "source_hash": source_hash, "drive_folder_id": root, "status": "processed",
    })[0]
    rows = [
        {"analysis_id": analysis["id"], "file_kind": "excel_original", "file_name": original_name, "drive_file_id": excel_id, "drive_web_url": excel_url},
        {"analysis_id": analysis["id"], "file_kind": "shp", "file_name": f"{farm_name}_limite.shp", "drive_file_id": shp_id, "drive_web_url": shp_url},
        {"analysis_id": analysis["id"], "file_kind": "reanálise_2026", "file_name": f"{farm_name}_reanálise_{year}.xlsx", "drive_file_id": reanalysis_id, "drive_web_url": reanalysis_url},
    ]
    supabase_request("analysis_files?on_conflict=analysis_id,file_kind", "POST", rows)
    return {"analysis_id": analysis["id"], "drive_folder_id": root}
