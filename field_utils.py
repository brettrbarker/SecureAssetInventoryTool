"""
Shared helpers for building field lists across windows.

Provides:
- compute_db_fields_from_template: fields limited to current template and not excluded
- compute_dropdown_fields: subset based on config.dropdown_fields
- compute_date_fields: subset of known date-like DB columns
"""

from __future__ import annotations

import os
import csv
from typing import List, Dict, Any


def compute_db_fields_from_template(db, config) -> List[Dict[str, str]]:
    """Build [{ 'db_name', 'display_name' }] limited to template headers and excluding config.excluded_fields.

    - Uses AssetDatabase.get_dynamic_column_mapping(template_path) to map headers -> db columns
    - Verifies the DB column exists in the assets table and is not a system column
    - Falls back to DB columns (converted to headers) when template isn't available
    """
    template_path = getattr(config, 'default_template_path', None)
    headers: List[str] = []
    if template_path and os.path.exists(template_path):
        try:
            with open(template_path, 'r', newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                headers = next(reader)
        except Exception:
            headers = []

    column_mapping = db.get_dynamic_column_mapping(template_path) if template_path else {}
    table_columns = set(db.get_table_columns())
    system_columns = {'id', 'created_date', 'modified_date', 'created_by', 'modified_by', 'is_deleted'}
    excluded_headers = set(getattr(config, 'excluded_fields', []) or [])

    fields: List[Dict[str, str]] = []
    for header in headers:
        if not header or not header.strip():
            continue
        if header in excluded_headers:
            continue
        db_col = column_mapping.get(header)
        if not db_col:
            continue
        if db_col in system_columns:
            continue
        if db_col not in table_columns:
            continue
        fields.append({'db_name': db_col, 'display_name': header})

    if fields:
        return fields

    # Fallback: derive from DB columns
    try:
        all_columns = db.get_table_columns()
        editable_columns = [col for col in all_columns if col not in system_columns]
        for col in editable_columns:
            readable = db._column_to_header(col)  # noqa: SLF001 (private access in local project)
            if readable in excluded_headers:
                continue
            fields.append({'db_name': col, 'display_name': readable})
    except Exception:
        pass

    return fields


def compute_dropdown_fields(db_fields: List[Dict[str, str]], config) -> List[Dict[str, str]]:
    """Return fields whose display names are listed in config.dropdown_fields.

    Only fields present in db_fields are considered (already template/exclusion filtered).
    """
    configured_dropdown_headers = set(getattr(config, 'dropdown_fields', []) or [])
    return [f for f in db_fields if f['display_name'] in configured_dropdown_headers]


def compute_date_fields(db_fields: List[Dict[str, str]]) -> List[Dict[str, str]]:
    """Return fields recognized as dates among the provided db_fields by DB column name."""
    date_field_names = {'audit_date', 'entry_date', 'created_date', 'modified_date'}
    return [f for f in db_fields if f['db_name'].lower() in date_field_names]
