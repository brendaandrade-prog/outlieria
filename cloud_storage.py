from __future__ import annotations

import hashlib
import io
import math
import os
import re
from datetime import datetime, timezone
from typing import Any, Iterable

import requests


EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SHAPEFILE_MIME = "application/octet-stream"


def setting(name: str) -> str:
    value = os.getenv(name, "")
    if value:
        return value
    try:
        import streamlit as st

        return str(st.secrets.get(name, ""))
    except Exception:
        return ""


def database_configured() -> bool:
    """Indica se o índice e os resultados do Supabase podem ser acessados."""
    return bool(setting("SUPABASE_URL") and setting("SUPABASE_SERVICE_ROLE_KEY"))


def configured() -> bool:
    """Indica se Supabase e Google Drive estão prontos para uma sincronização completa."""
    return bool(
        database_configured()
        and setting("GOOGLE_CLIENT_ID")
        and setting("GOOGLE_CLIENT_SECRET")
        and setting("GOOGLE_REFRESH_TOKEN")
    )


def missing_settings() -> list[str]:
    """Lista apenas nomes ausentes; nunca expõe valores de Secrets."""
    required = [
        "SUPABASE_URL",
        "SUPABASE_SERVICE_ROLE_KEY",
        "GOOGLE_CLIENT_ID",
        "GOOGLE_CLIENT_SECRET",
        "GOOGLE_REFRESH_TOKEN",
    ]
    return [name for name in required if not setting(name)]


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _clean_value(value: Any) -> Any:
    """Converte valores pandas/numpy e NaN para tipos aceitos pelo PostgREST."""
    if value is None:
        return None
    try:
        if isinstance(value, float) and math.isnan(value):
            return None
    except TypeError:
        pass
    if hasattr(value, "item"):
        try:
            value = value.item()
        except Exception:
            pass
    try:
        if value != value:  # noqa: PLR0124 - cobre alguns escalares numpy
            return None
    except Exception:
        pass
    if isinstance(value, (list, tuple)):
        return [_clean_value(item) for item in value]
    if isinstance(value, dict):
        return {str(key): _clean_value(item) for key, item in value.items()}
    return value


def _chunks(items: list[dict[str, Any]], size: int = 200) -> Iterable[list[dict[str, Any]]]:
    for start in range(0, len(items), size):
        yield items[start : start + size]


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
    response = requests.request(method, f"{base}/rest/v1/{path}", headers=headers, params=params, json=payload, timeout=60)
    try:
        response.raise_for_status()
    except requests.HTTPError as exc:
        detail = response.text.strip()
        if detail:
            raise requests.HTTPError(f"{exc} — resposta Supabase: {detail[:800]}", response=response) from exc
        raise
    return response.json() if response.content else None


def _is_missing_column_error(exc: Exception) -> bool:
    detail = str(exc).lower()
    return any(
        marker in detail
        for marker in (
            "could not find",
            "schema cache",
            "column .* does not exist",
            "undefined column",
            "42703",
            "pgrst204",
        )
    )


def _missing_column_name(exc: Exception) -> str | None:
    match = re.search(r"Could not find the '([^']+)' column", str(exc), flags=re.IGNORECASE)
    return match.group(1) if match else None


def _analysis_files(analysis_id: str) -> list[dict[str, Any]]:
    return (
        supabase_request(
            "analysis_files",
            params={
                "analysis_id": f"eq.{analysis_id}",
                "select": "file_kind,file_name,drive_file_id,drive_web_url,created_at",
                "order": "file_kind.asc",
            },
        )
        or []
    )


def _find_existing_analysis(*, source_hash: str, client_name: str, farm_name: str, year: int, original_name: str, shp_name: str) -> tuple[dict[str, Any] | None, list[dict[str, Any]]]:
    """Localiza a análise exata ou a mesma fonte lógica com conteúdo atualizado."""
    exact = supabase_request(
        "analyses",
        params={
            "source_hash": f"eq.{source_hash}",
            "select": "*",
            "limit": "1",
        },
    ) or []
    if exact:
        analysis = exact[0]
        return analysis, _analysis_files(str(analysis["id"]))

    candidates = supabase_request(
        "analyses",
        params={
            "client_name": f"eq.{client_name}",
            "farm_name": f"eq.{farm_name}",
            "analysis_year": f"eq.{year}",
            "select": "*",
            "order": "updated_at.desc,inserted_at.desc",
            "limit": "20",
        },
    ) or []
    for candidate in candidates:
        files = _analysis_files(str(candidate["id"]))
        names = {str(row.get("file_name") or "") for row in files}
        if original_name in names or shp_name in names:
            return candidate, files
    # Projetos antigos podem ter o resumo gravado sem os arquivos indexados.
    # Se houver somente um resumo para a mesma fazenda/ano, ele é a fonte lógica
    # que deve ser atualizada pelo novo envio.
    if len(candidates) == 1:
        candidate = candidates[0]
        return candidate, _analysis_files(str(candidate["id"]))
    return None, []


def _upsert_analysis_summary(payload: dict[str, Any], existing: dict[str, Any] | None) -> dict[str, Any]:
    if existing:
        path = f"analyses?id=eq.{existing['id']}"
        try:
            response = supabase_request(path, "PATCH", payload)
        except requests.HTTPError as exc:
            if "lof_min_neighbors" not in str(exc).lower():
                raise
            response = supabase_request(path, "PATCH", {key: value for key, value in payload.items() if key != "lof_min_neighbors"})
        return (response or [None])[0] or {**existing, **payload}

    try:
        response = supabase_request("analyses?on_conflict=source_hash", "POST", payload)
    except requests.HTTPError as exc:
        if "lof_min_neighbors" not in str(exc).lower():
            raise
        response = supabase_request(
            "analyses?on_conflict=source_hash",
            "POST",
            {key: value for key, value in payload.items() if key != "lof_min_neighbors"},
        )
    if not response:
        raise RuntimeError("O Supabase não retornou o registro da análise após o upsert.")
    return response[0]


def _point_identity(row: dict[str, Any]) -> tuple[Any, ...]:
    def numeric(value: Any) -> Any:
        try:
            return round(float(value), 8)
        except (TypeError, ValueError):
            return None

    return (
        _clean_value(row.get("source_row")),
        str(row.get("point_id") or ""),
        int(row.get("year") or 0),
        str(row.get("crop") or ""),
        str(row.get("depth") or ""),
        numeric(row.get("longitude")),
        numeric(row.get("latitude")),
    )


def _point_key(row: dict[str, Any]) -> str:
    identity = _point_identity(row)
    return "|".join("" if value is None else str(value) for value in identity)


def _insert_points(analysis_id: str, result_rows: list[dict[str, Any]]) -> dict[tuple[Any, ...], str]:
    points: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()
    for row in result_rows:
        identity = _point_identity(row)
        if identity in seen:
            continue
        seen.add(identity)
        points.append(
            {
                "analysis_id": analysis_id,
                "point_key": _point_key(row),
                "point_id": _clean_value(row.get("point_id")),
                "sample_index": _clean_value(row.get("sample_index")),
                "source_row": _clean_value(row.get("source_row")),
                "year": _clean_value(row.get("year")),
                "crop": _clean_value(row.get("crop")),
                "depth": _clean_value(row.get("depth")),
                "longitude": _clean_value(row.get("longitude")),
                "latitude": _clean_value(row.get("latitude")),
                "municipality": _clean_value(row.get("municipality")),
                "state": _clean_value(row.get("state")),
                "polygon_index": _clean_value(row.get("polygon_index")),
            }
        )
    if not points:
        return {}

    inserted: list[dict[str, Any]] = []
    try:
        for chunk in _chunks(points):
            inserted.extend(supabase_request("analysis_points", "POST", chunk) or [])
    except requests.HTTPError as exc:
        if not _is_missing_column_error(exc):
            raise
        # Compatibilidade com tabelas criadas antes de point_key/sample_index.
        supabase_request(f"analysis_points?analysis_id=eq.{analysis_id}", "DELETE")
        legacy_points = [{key: value for key, value in row.items() if key not in {"point_key", "sample_index"}} for row in points]
        inserted = []
        for chunk in _chunks(legacy_points):
            inserted.extend(supabase_request("analysis_points", "POST", chunk) or [])

    point_ids: dict[tuple[Any, ...], str] = {}
    for row in inserted:
        key = row.get("point_key")
        if key is not None and row.get("id"):
            for original in points:
                if original.get("point_key") == key:
                    point_ids[_point_identity(original)] = row["id"]
                    break
        elif row.get("id"):
            identity = _point_identity(row)
            point_ids[identity] = row["id"]
    # O PostgREST conserva a ordem do lote como último recurso para bases antigas.
    if len(point_ids) < len(inserted):
        for original, returned in zip(points, inserted):
            if returned.get("id"):
                point_ids.setdefault(_point_identity(original), returned["id"])
    return point_ids


def _result_payload(analysis_id: str, point_id: str, row: dict[str, Any]) -> dict[str, Any]:
    keys = (
        "element", "value", "historical_mean", "historical_median", "q1", "q3", "iqr",
        "lower_fence", "upper_fence", "statistical_status", "direction", "ml_status", "ml_score",
        "ml_neighbor_count", "ml_k_effective", "ml_radius_m", "ml_rule", "ml_mean_distance_m",
        "ml_regional_median", "ml_reference_scope", "ml_historical_percentile",
        "ml_historical_distribution_shift", "ml_current_cohort_shift", "reg_status", "reg_prediction",
        "reg_residual", "reg_score", "reg_cycles", "reg_rule", "ph_smp_status", "display_status",
        "sample_index", "final_status",
    )
    payload = {"analysis_id": analysis_id, "point_id_ref": point_id}
    payload.update({key: _clean_value(row.get(key)) for key in keys})
    return payload


def _insert_results(analysis_id: str, result_rows: list[dict[str, Any]], point_ids: dict[tuple[Any, ...], str]) -> None:
    rows = []
    for row in result_rows:
        point_id = point_ids.get(_point_identity(row))
        if not point_id or not row.get("element"):
            continue
        rows.append(_result_payload(analysis_id, point_id, row))
    if not rows:
        return
    try:
        for chunk in _chunks(rows):
            supabase_request("analysis_results", "POST", chunk)
    except requests.HTTPError as exc:
        if not _is_missing_column_error(exc):
            raise
        # Instalações antigas podem estar vários commits atrás do schema.
        # Remove exatamente a coluna reportada pelo PostgREST e tenta de novo
        # até chegar ao conjunto de campos suportado pela tabela real.
        compatible_rows = [dict(row) for row in rows]
        missing_column = _missing_column_name(exc)
        if missing_column:
            for row in compatible_rows:
                row.pop(missing_column, None)
        else:
            compatible_rows = [
                {key: value for key, value in row.items() if key in {
                    "analysis_id", "point_id_ref", "element", "value", "historical_mean", "q1", "q3", "iqr",
                    "lower_fence", "upper_fence", "statistical_status", "direction", "ml_status", "ml_score",
                    "ml_neighbor_count", "ml_radius_m", "ml_rule", "ml_mean_distance_m", "final_status",
                }}
                for row in rows
            ]
        for _ in range(30):
            supabase_request(f"analysis_results?analysis_id=eq.{analysis_id}", "DELETE")
            try:
                for chunk in _chunks(compatible_rows):
                    supabase_request("analysis_results", "POST", chunk)
                return
            except requests.HTTPError as retry_exc:
                if not _is_missing_column_error(retry_exc):
                    raise
                missing_column = _missing_column_name(retry_exc)
                if not missing_column or not any(missing_column in row for row in compatible_rows):
                    raise
                for row in compatible_rows:
                    row.pop(missing_column, None)
        raise RuntimeError("O schema Supabase possui mais de 30 colunas incompatíveis em analysis_results.")


def _replace_persisted_rows(analysis_id: str, result: Any) -> None:
    """Substitui os filhos da análise, evitando duplicatas em cada reenvio."""
    supabase_request(f"analysis_results?analysis_id=eq.{analysis_id}", "DELETE")
    supabase_request(f"analysis_points?analysis_id=eq.{analysis_id}", "DELETE")
    result_rows = result.to_dict(orient="records") if result is not None and hasattr(result, "to_dict") else list(result or [])
    result_rows = [{str(key): _clean_value(value) for key, value in row.items()} for row in result_rows]
    point_ids = _insert_points(analysis_id, result_rows)
    _insert_results(analysis_id, result_rows, point_ids)


def persist_analysis(
    *,
    client_name: str,
    farm_name: str,
    year: int,
    radius_m: float,
    original: bytes,
    shp_bytes: bytes,
    reanalysis: bytes,
    original_name: str,
    shp_name: str | None = None,
    area_ha: float,
    point_count: int,
    outlier_count: int,
    attention_count: int,
    result: Any = None,
    lof_min_neighbors: int = 3,
) -> dict[str, Any] | None:
    """Sincroniza a fonte, resumo e resultados; reenvios atualizam o mesmo registro."""
    if not database_configured():
        return None

    shp_name = shp_name or f"{farm_name}_limite.shp"
    source_hash = sha256(original + shp_bytes)
    existing, existing_files = _find_existing_analysis(
        source_hash=source_hash,
        client_name=client_name,
        farm_name=farm_name,
        year=year,
        original_name=original_name,
        shp_name=shp_name,
    )
    root = (existing or {}).get("drive_folder_id")
    file_urls: dict[str, str] = {}

    if configured():
        service = drive_service()
        root = root or ensure_folder(service, "Outlieria")
        originals = ensure_folder(service, "01_Fontes_Originais", root)
        reanalyses = ensure_folder(service, "02_Reanalises_Individuais", root)
        existing_by_kind = {row.get("file_kind"): row for row in existing_files}
        if existing_by_kind.get("excel_original", {}).get("drive_file_id"):
            excel_id, excel_url = update_bytes(service, existing_by_kind["excel_original"]["drive_file_id"], original, EXCEL_MIME)
        else:
            excel_id, excel_url = upload_bytes(service, original_name, original, EXCEL_MIME, originals)
        if existing_by_kind.get("shp", {}).get("drive_file_id"):
            shp_id, shp_url = update_bytes(service, existing_by_kind["shp"]["drive_file_id"], shp_bytes, SHAPEFILE_MIME)
        else:
            shp_id, shp_url = upload_bytes(service, shp_name, shp_bytes, SHAPEFILE_MIME, originals)
        if existing_by_kind.get("reanálise_2026", {}).get("drive_file_id"):
            reanalysis_id, reanalysis_url = update_bytes(service, existing_by_kind["reanálise_2026"]["drive_file_id"], reanalysis, EXCEL_MIME)
        else:
            reanalysis_id, reanalysis_url = upload_bytes(service, f"{farm_name}_reanálise_{year}.xlsx", reanalysis, EXCEL_MIME, reanalyses)
        file_rows = [
            {"analysis_id": existing["id"] if existing else None, "file_kind": "excel_original", "file_name": original_name, "drive_file_id": excel_id, "drive_web_url": excel_url},
            {"analysis_id": existing["id"] if existing else None, "file_kind": "shp", "file_name": shp_name, "drive_file_id": shp_id, "drive_web_url": shp_url},
            {"analysis_id": existing["id"] if existing else None, "file_kind": "reanálise_2026", "file_name": f"{farm_name}_reanálise_{year}.xlsx", "drive_file_id": reanalysis_id, "drive_web_url": reanalysis_url},
        ]
        file_urls = {row["file_kind"]: row["drive_web_url"] for row in file_rows}
    else:
        file_rows = []

    analysis_payload = {
        "client_name": client_name,
        "farm_name": farm_name,
        "analysis_year": year,
        "radius_m": round(radius_m, 2),
        "lof_min_neighbors": int(lof_min_neighbors),
        "area_2026_ha": round(area_ha, 2),
        "point_count": point_count,
        "outlier_count": outlier_count,
        "attention_count": attention_count,
        "source_hash": source_hash,
        "drive_folder_id": root,
        "status": "processed",
        "updated_at": _now_iso(),
    }
    analysis = _upsert_analysis_summary(analysis_payload, existing)
    analysis_id = str(analysis["id"])

    if file_rows:
        for row in file_rows:
            row["analysis_id"] = analysis_id
        supabase_request("analysis_files?on_conflict=analysis_id,file_kind", "POST", file_rows)

    if result is not None:
        _replace_persisted_rows(analysis_id, result)

    return {
        "analysis_id": analysis_id,
        "drive_folder_id": root,
        "source_hash": source_hash,
        "replaced": bool(existing),
        "file_urls": file_urls,
    }


def delete_analysis(analysis_id: str) -> dict[str, Any]:
    """Exclui arquivos da análise no Drive e o registro Supabase em cascata."""
    if not database_configured():
        raise RuntimeError("Supabase não está configurado; a análise não foi excluída do banco.")
    analysis_id = str(analysis_id)
    files = _analysis_files(analysis_id)
    drive_deleted = 0
    drive_errors: list[str] = []
    if configured() and files:
        service = drive_service()
        for row in files:
            file_id = row.get("drive_file_id")
            if not file_id:
                continue
            try:
                service.files().delete(fileId=file_id).execute()
                drive_deleted += 1
            except Exception as exc:
                # O arquivo pode já ter sido removido manualmente no Drive.
                detail = str(exc)
                if "404" not in detail and "not found" not in detail.lower():
                    drive_errors.append(f"{row.get('file_name') or file_id}: {detail}")
        if drive_errors:
            raise RuntimeError("Não foi possível excluir todos os arquivos do Google Drive: " + " | ".join(drive_errors))

    supabase_request(f"analyses?id=eq.{analysis_id}", "DELETE")
    return {
        "analysis_id": analysis_id,
        "drive_deleted": drive_deleted,
        "drive_configured": configured(),
        "drive_files_pending": bool(files) and not configured(),
    }


def _result_rows_for_analysis(analysis: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    analysis_id = str(analysis["id"])
    points = supabase_request(
        "analysis_points",
        params={
            "analysis_id": f"eq.{analysis_id}",
            "select": "*",
            "order": "source_row.asc,point_id.asc",
        },
    ) or []
    results = supabase_request(
        "analysis_results",
        params={"analysis_id": f"eq.{analysis_id}", "select": "*", "order": "id.asc"},
    ) or []
    points_by_id = {str(point.get("id")): point for point in points if point.get("id")}
    rows: list[dict[str, Any]] = []
    for result in results:
        point = points_by_id.get(str(result.get("point_id_ref")), {})
        year = point.get("year") or analysis.get("analysis_year") or 2026
        if str(year) != str(analysis.get("analysis_year") or 2026):
            continue
        row = {
            "analysis_id": analysis_id,
            "client": analysis.get("client_name", ""),
            "farm": analysis.get("farm_name", ""),
            "year": str(year),
            "sample_index": result.get("sample_index") if result.get("sample_index") is not None else point.get("sample_index", point.get("source_row")),
            "point_id": point.get("point_id", ""),
            "source_row": point.get("source_row"),
            "sheet": str(year),
            "crop": point.get("crop", ""),
            "depth": point.get("depth", ""),
            "longitude": point.get("longitude"),
            "latitude": point.get("latitude"),
            "municipality": point.get("municipality", "Não localizado"),
            "state": point.get("state", "Não localizado"),
            "polygon_index": point.get("polygon_index", -1),
        }
        for key, value in result.items():
            if key not in {"id", "analysis_id", "point_id_ref", "created_at"}:
                row[key] = value
        rows.append(row)
    return rows, points


def load_persisted_analyses(year: int = 2026) -> list[dict[str, Any]]:
    """Carrega fontes e resultados já gravados para a inicialização da interface."""
    if not database_configured():
        return []
    analyses = supabase_request(
        "analyses",
        params={
            "analysis_year": f"eq.{year}",
            "select": "*",
            "order": "updated_at.desc,inserted_at.desc",
        },
    ) or []
    loaded: list[dict[str, Any]] = []
    for analysis in analyses:
        result_rows, _ = _result_rows_for_analysis(analysis)
        files = _analysis_files(str(analysis["id"]))
        file_urls = {str(row.get("file_kind")): row.get("drive_web_url") or "" for row in files}
        original_file = next((row for row in files if row.get("file_kind") == "excel_original"), {})
        shp_file = next((row for row in files if row.get("file_kind") == "shp"), {})
        loaded.append(
            {
                "analysis_id": str(analysis["id"]),
                "source_hash": analysis.get("source_hash", ""),
                "name": original_file.get("file_name") or f"{analysis.get('farm_name', 'fonte')}_{year}.xlsx",
                "shp_name": shp_file.get("file_name") or f"{analysis.get('farm_name', 'fonte')}_limite.shp",
                "client_name": analysis.get("client_name", ""),
                "farm_name": analysis.get("farm_name", ""),
                "radius_m": float(analysis.get("radius_m") or 300),
                "area_ha": float(analysis.get("area_2026_ha") or 0),
                "result_rows": result_rows,
                "file_urls": file_urls,
                "updated_at": analysis.get("updated_at") or analysis.get("inserted_at"),
                "point_count": analysis.get("point_count") or len({row.get("source_row") for row in result_rows}),
                "outlier_count": analysis.get("outlier_count") or 0,
                "attention_count": analysis.get("attention_count") or 0,
            }
        )
    return loaded
