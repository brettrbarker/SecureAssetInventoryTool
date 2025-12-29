"""Local web UI entrypoint using FastAPI.
Run with `python web_app.py` and visit http://127.0.0.1:8000.
"""
from __future__ import annotations

import csv
import json
import os
import shutil
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any, Dict, List, Optional, Tuple

from fastapi import FastAPI, HTTPException, UploadFile, File, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from config_manager import ConfigManager
from database_service import DatabaseService
from asset_database import AssetDatabase

config_manager = ConfigManager()
config = config_manager.get_config()

db_service = DatabaseService()
db: AssetDatabase = db_service.get_database_instance()

webui_dir = Path(__file__).parent / "webui"

app = FastAPI(title="Secure Asset Inventory Tool - Web UI", version="1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://127.0.0.1:8000", "http://localhost:8000"],
    allow_credentials=True,
    allow_methods=["*"]
)

app.mount("/static", StaticFiles(directory=webui_dir), name="static")


class AddAssetRequest(BaseModel):
    fields: Dict[str, Any] = Field(..., description="Template header -> value mapping")


class AssetResponse(BaseModel):
    asset_id: int
    asset: Dict[str, Any]


class UpdateAssetRequest(BaseModel):
    fields: Dict[str, Any]


class BulkOperation(BaseModel):
    field: str
    operation: str = Field("replace", description="replace|append")
    value: Any = None


class BulkUpdateRequest(BaseModel):
    filters: Optional[Dict[str, Any]] = None
    ids: Optional[List[int]] = None
    operations: List[BulkOperation]
    limit: int = 500


class SavedSearch(BaseModel):
    name: str
    filters: Dict[str, Any]


class SettingsUpdate(BaseModel):
    theme: Optional[str] = None
    default_template_path: Optional[str] = None
    output_directory: Optional[str] = None
    database_path: Optional[str] = None
    dropdown_fields: Optional[List[str]] = None
    required_fields: Optional[List[str]] = None
    excluded_fields: Optional[List[str]] = None
    unique_fields: Optional[List[str]] = None
    monitor_primary_fields: Optional[List[str]] = None
    monitor_secondary_fields: Optional[List[str]] = None
    monitor_tertiary_fields: Optional[List[str]] = None
    label_output_fields: Optional[List[str]] = None
    hmr_fields: Optional[List[str]] = None
    destruction_report_fields: Optional[List[str]] = None
    bulk_update_presets: Optional[Dict[str, Any]] = None
    saved_searches: Optional[Dict[str, Any]] = None


@app.get("/", include_in_schema=False)
async def serve_index() -> FileResponse:
    index_path = webui_dir / "index.html"
    if not index_path.exists():
        raise HTTPException(status_code=500, detail="Web UI assets missing. Build webui first.")
    return FileResponse(index_path)


@app.get("/api/health")
def health() -> Dict[str, str]:
    return {"status": "ok"}


def _get_template_path() -> Optional[str]:
    template_path = config_manager.get_template_path()
    return template_path if template_path else None


def _get_template_headers() -> List[str]:
    template_path = _get_template_path()
    if not template_path or not Path(template_path).exists():
        return []
    try:
        with open(template_path, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            return next(reader, []) or []
    except Exception:
        return []


def _column_mapping() -> Dict[str, str]:
    template_path = _get_template_path()
    return db.get_dynamic_column_mapping(template_path) if template_path else {}


def _to_db_payload(fields: Dict[str, Any]) -> Dict[str, Any]:
    mapping = _column_mapping()
    payload: Dict[str, Any] = {}
    for header, value in fields.items():
        db_column = mapping.get(header)
        if db_column is None:
            continue
        payload[db_column] = value
    return payload


def _refresh_database_binding(new_db_path: Optional[str]) -> None:
    """Rebind the shared DatabaseService/AssetDatabase to a new path if provided."""
    global db, db_service
    if new_db_path:
        db = AssetDatabase(new_db_path)
        db_service.db = db


def _build_filters(mapping: Dict[str, str], filters: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    if not filters:
        return {}
    resolved: Dict[str, Any] = {}
    for k, v in filters.items():
        if v in (None, ""):
            continue
        db_field = mapping.get(k) or k
        resolved[db_field] = v
    return resolved


def _query_assets(
    filters: Dict[str, Any],
    sort_by: str = "modified_date",
    sort_dir: str = "desc",
    limit: int = 200,
    offset: int = 0,
) -> Tuple[List[Dict[str, Any]], int]:
    available_columns = set(db.get_table_columns())

    # Only allow sorting on known columns
    sort_column = sort_by if sort_by in available_columns else "modified_date"
    direction = "DESC" if sort_dir.lower() == "desc" else "ASC"

    # Build WHERE clause
    where = []
    params: List[Any] = []
    for field, value in filters.items():
        if field not in available_columns:
            continue
        where.append(f"{field} LIKE ?" if isinstance(value, str) else f"{field} = ?")
        params.append(f"%{value}%" if isinstance(value, str) else value)
    where_sql = " AND ".join(where) if where else "1=1"

    # Count query
    with db.get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(f"SELECT COUNT(*) FROM assets WHERE {where_sql}", params)
        total = cursor.fetchone()[0]

        cursor.execute(
            f"SELECT * FROM assets WHERE {where_sql} ORDER BY {sort_column} {direction} LIMIT ? OFFSET ?",
            params + [limit, offset],
        )
        rows = cursor.fetchall()
        columns = [c[0] for c in cursor.description]
        items = [dict(zip(columns, row)) for row in rows]
    return items, total


def _map_fields_to_db(fields: Dict[str, Any]) -> Dict[str, Any]:
    mapping = _column_mapping()
    payload: Dict[str, Any] = {}
    for header, value in fields.items():
        db_col = mapping.get(header) or header
        payload[db_col] = value
    return payload


def _apply_bulk_operations(asset: Dict[str, Any], operations: List[BulkOperation], template_path: Optional[str]) -> Dict[str, Any]:
    updates: Dict[str, Any] = {}
    mapping = _column_mapping()
    for op in operations:
        db_col = mapping.get(op.field) or op.field
        if not db_col:
            continue
        value = op.value
        if value == "current_date":
            value = datetime.now().strftime("%m/%d/%Y")

        operation = (op.operation or "replace").lower()
        if operation not in {"replace", "append"}:
            operation = "replace"

        if operation == "replace":
            updates[db_col] = value
        else:
            existing = asset.get(db_col) or ""
            if existing:
                if db.should_field_be_multiline(op.field, template_path):
                    updates[db_col] = f"{existing}\n{value}"
                else:
                    updates[db_col] = f"{existing} {value}".strip()
            else:
                updates[db_col] = value
    return updates


@app.get("/api/config")
def get_config() -> Dict[str, Any]:
    cfg = config_manager.get_config()
    return {
        "template_path": cfg.default_template_path,
        "database_path": cfg.database_path,
        "dropdown_fields": cfg.dropdown_fields,
        "required_fields": cfg.required_fields,
        "unique_fields": cfg.unique_fields,
    }


@app.get("/api/settings")
def get_settings() -> Dict[str, Any]:
    cfg = config_manager.get_config()
    return {
        "theme": cfg.theme,
        "default_template_path": cfg.default_template_path,
        "output_directory": cfg.output_directory,
        "database_path": cfg.database_path,
        "dropdown_fields": cfg.dropdown_fields,
        "required_fields": cfg.required_fields,
        "excluded_fields": cfg.excluded_fields,
        "unique_fields": cfg.unique_fields,
        "monitor_primary_fields": cfg.monitor_primary_fields,
        "monitor_secondary_fields": cfg.monitor_secondary_fields,
        "monitor_tertiary_fields": cfg.monitor_tertiary_fields,
        "label_output_fields": cfg.label_output_fields,
        "hmr_fields": cfg.hmr_fields,
        "destruction_report_fields": cfg.destruction_report_fields,
        "bulk_update_presets": cfg.bulk_update_presets,
        "saved_searches": cfg.saved_searches,
        "template_headers": _get_template_headers(),
    }


@app.put("/api/settings")
def update_settings(payload: SettingsUpdate) -> Dict[str, Any]:
    cfg_before = config_manager.get_config()
    updates: Dict[str, Any] = {}

    # Apply any provided fields
    for field_name, value in payload.model_dump(exclude_none=True).items():
        updates[field_name] = value

    if not updates:
        return {"updated": False, "message": "No changes supplied"}

    # Apply updates to config
    config_manager.update_config(**updates)
    updated_cfg = config_manager.get_config()

    # Ensure directories for new paths
    config_manager.ensure_directories()

    # Refresh DB binding if path changed
    if updates.get("database_path") and updates["database_path"] != cfg_before.database_path:
        _refresh_database_binding(updates["database_path"])

    # If template changed and exists, attempt schema update
    template_path = updated_cfg.default_template_path
    if updates.get("default_template_path") and Path(template_path).exists():
        try:
            db.update_schema_for_template(template_path)
        except Exception:
            pass

    config_manager.save_config()
    return {"updated": True, "settings": get_settings()}


@app.get("/api/fields")
def get_fields() -> Dict[str, Any]:
    template_path = _get_template_path()
    metadata = db.get_field_metadata(template_path) if template_path else {}
    mapping = _column_mapping()
    cfg = config_manager.get_config()
    dropdown_values = db_service.get_dropdown_values(template_path, cfg.dropdown_fields)
    return {
        "field_metadata": metadata,
        "column_mapping": mapping,
        "dropdown_fields": cfg.dropdown_fields,
        "required_fields": cfg.required_fields,
        "unique_fields": cfg.unique_fields,
        "dropdown_values": dropdown_values,
    }


@app.get("/api/assets")
def list_assets(
    query: Optional[str] = None,
    field: Optional[str] = None,
    limit: int = 200,
    offset: int = 0,
    sort_by: Optional[str] = "modified_date",
    sort_dir: Optional[str] = "desc",
    filters: Optional[str] = Query(None, description="JSON object of additional filters"),
    include_deleted: bool = False,
) -> Dict[str, Any]:
    if limit <= 0 or limit > 2000:
        raise HTTPException(status_code=400, detail="Limit must be between 1 and 2000.")
    if offset < 0:
        offset = 0

    mapping = _column_mapping()
    db_field = mapping.get(field) or field if field else None

    parsed_filters: Dict[str, Any] = {}
    if filters:
        try:
            parsed_filters = json.loads(filters)
        except json.JSONDecodeError:
            raise HTTPException(status_code=400, detail="Invalid filters JSON")

    resolved_filters = _build_filters(mapping, parsed_filters)
    resolved_filters["is_deleted"] = 0 if not include_deleted else resolved_filters.get("is_deleted", None)
    if db_field and query:
        resolved_filters[db_field] = query
    # Clean up None entries
    resolved_filters = {k: v for k, v in resolved_filters.items() if v is not None}

    items, total = _query_assets(resolved_filters, sort_by or "modified_date", sort_dir or "desc", limit, offset)
    return {"items": items, "count": len(items), "total": total}


@app.get("/api/assets/{asset_id}")
def get_asset(asset_id: int) -> Dict[str, Any]:
    asset = db.get_asset_by_id(asset_id)
    if not asset:
        raise HTTPException(status_code=404, detail="Asset not found")
    return asset


@app.post("/api/assets", response_model=AssetResponse)
def create_asset(payload: AddAssetRequest) -> AssetResponse:
    template_path = _get_template_path()
    if not template_path:
        raise HTTPException(status_code=400, detail="Template path is not configured.")

    cfg = config_manager.get_config()
    db_payload = _to_db_payload(payload.fields)

    conflicts = db.check_unique_field_conflicts(db_payload, cfg.unique_fields, template_path)
    if conflicts:
        raise HTTPException(status_code=409, detail={"conflicts": conflicts})

    asset_id = db_service.add_asset_from_form(payload.fields, template_path)
    if asset_id is None:
        raise HTTPException(status_code=500, detail="Unable to create asset")

    created_asset = db.get_asset_by_id(asset_id) or {}
    return AssetResponse(asset_id=asset_id, asset=created_asset)


@app.put("/api/assets/{asset_id}")
def update_asset(asset_id: int, payload: UpdateAssetRequest) -> Dict[str, Any]:
    template_path = _get_template_path()
    db_payload = _map_fields_to_db(payload.fields)
    success = db.update_asset(asset_id, db_payload)
    if not success:
        raise HTTPException(status_code=404, detail="Asset not found or not updated")
    updated = db.get_asset_by_id(asset_id)
    return {"updated": True, "asset": updated}


@app.delete("/api/assets/{asset_id}")
def delete_asset(asset_id: int) -> Dict[str, Any]:
    success = db.delete_asset(asset_id)
    if not success:
        raise HTTPException(status_code=404, detail="Asset not found")
    return {"deleted": True, "asset_id": asset_id}


@app.post("/api/assets/bulk-update")
def bulk_update(payload: BulkUpdateRequest) -> Dict[str, Any]:
    template_path = _get_template_path()
    mapping = _column_mapping()
    filters = _build_filters(mapping, payload.filters)
    filters = {k: v for k, v in filters.items() if v is not None}
    filters.setdefault("is_deleted", 0)

    # Determine target assets
    targets: List[Dict[str, Any]] = []
    if payload.ids:
        for aid in payload.ids:
            asset = db.get_asset_by_id(aid)
            if asset:
                targets.append(asset)
        if not targets:
            raise HTTPException(status_code=404, detail="No matching assets for provided ids")
    else:
        limit = min(max(payload.limit, 1), 2000)
        targets, _ = _query_assets(filters, limit=limit, offset=0)
        if not targets:
            raise HTTPException(status_code=404, detail="No assets match filters")

    updated_ids: List[int] = []
    for asset in targets:
        updates = _apply_bulk_operations(asset, payload.operations, template_path)
        if not updates:
            continue
        try:
            db.update_asset(asset["id"], updates)
            updated_ids.append(asset["id"])
        except Exception:
            continue

    return {"updated": len(updated_ids), "ids": updated_ids, "attempted": len(targets)}


@app.post("/api/assets/import")
async def import_assets(file: UploadFile = File(...)) -> Dict[str, Any]:
    # Save upload to temp file
    suffix = Path(file.filename or "upload.csv").suffix or ".csv"
    with NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name

    try:
        imported = db_service.import_assets_from_csv(tmp_path)
    finally:
        os.remove(tmp_path)

    return {"imported": imported}


@app.get("/api/assets/export")
def export_assets(
    mode: str = "template",
    query: Optional[str] = None,
    field: Optional[str] = None,
    filters: Optional[str] = Query(None, description="JSON filters"),
    limit: int = 2000,
) -> FileResponse:
    if limit <= 0 or limit > 10000:
        raise HTTPException(status_code=400, detail="Limit must be between 1 and 10000")

    mapping = _column_mapping()
    db_field = mapping.get(field) or field if field else None
    parsed_filters: Dict[str, Any] = {}
    if filters:
        try:
            parsed_filters = json.loads(filters)
        except json.JSONDecodeError:
            raise HTTPException(status_code=400, detail="Invalid filters JSON")

    resolved_filters = _build_filters(mapping, parsed_filters)
    resolved_filters.setdefault("is_deleted", 0)
    if db_field and query:
        resolved_filters[db_field] = query

    assets, _ = _query_assets(resolved_filters, limit=limit, offset=0)
    template_path = _get_template_path() if mode == "template" else None

    # Write CSV to temp file
    tmp = NamedTemporaryFile(delete=False, suffix=".csv")
    tmp_path = Path(tmp.name)
    tmp.close()

    try:
        success = db_service.export_assets_to_csv(assets, str(tmp_path), template_path)
        if not success:
            raise HTTPException(status_code=500, detail="Export failed")
    except Exception as exc:
        if tmp_path.exists():
            tmp_path.unlink()
        raise HTTPException(status_code=500, detail=str(exc))

    filename = f"assets_export_{mode}.csv"
    return FileResponse(path=str(tmp_path), filename=filename, media_type="text/csv")


@app.get("/api/saved-searches")
def get_saved_searches() -> Dict[str, Any]:
    cfg = config_manager.get_config()
    return cfg.saved_searches or {}


@app.post("/api/saved-searches")
def save_search(payload: SavedSearch) -> Dict[str, Any]:
    cfg = config_manager.get_config()
    saved = cfg.saved_searches or {}
    saved[payload.name] = payload.filters
    config_manager.update_config(saved_searches=saved)
    config_manager.save_config()
    return saved


@app.delete("/api/saved-searches/{name}")
def delete_saved_search(name: str) -> Dict[str, Any]:
    cfg = config_manager.get_config()
    saved = cfg.saved_searches or {}
    if name in saved:
        del saved[name]
        config_manager.update_config(saved_searches=saved)
        config_manager.save_config()
    return saved


@app.get("/api/monitor")
def monitor() -> Dict[str, Any]:
    stats = db_service.get_database_statistics()
    recent_modified = db_service.get_recently_modified_assets(days=7, exclude_new=False)
    recent_added = db_service.get_recently_added_assets(days=7)
    return {
        "stats": stats,
        "recent_modified": recent_modified,
        "recent_added": recent_added,
    }


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "web_app:app",
        host="127.0.0.1",
        port=8000,
        reload=False,
        log_level="info",
    )
