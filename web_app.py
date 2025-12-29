"""Local web UI entrypoint using FastAPI.
Run with `python web_app.py` and visit http://127.0.0.1:8000.
"""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, HTTPException
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
def list_assets(query: Optional[str] = None, field: Optional[str] = None, limit: int = 200) -> Dict[str, Any]:
    if limit <= 0 or limit > 2000:
        raise HTTPException(status_code=400, detail="Limit must be between 1 and 2000.")

    mapping = _column_mapping()
    db_field = None
    if field:
        db_field = mapping.get(field) or field

    filters: Dict[str, Any] = {"is_deleted": 0}
    if query and db_field:
        filters[db_field] = query

    assets = db.search_assets(filters, limit=limit)
    return {"items": assets, "count": len(assets)}


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


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "web_app:app",
        host="127.0.0.1",
        port=8000,
        reload=False,
        log_level="info",
    )
