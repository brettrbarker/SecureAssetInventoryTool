"""CLI to export audit log to CSV.

Usage examples:
    python export_audit_log.py --output assets/output_files/audit_log.csv
    python export_audit_log.py --action UPDATE --date-from 2026-01-01T00:00:00 --output updates.csv
"""
import argparse
from typing import Dict, Any, Optional

from database_service import database_service


def parse_args():
    p = argparse.ArgumentParser(description="Export audit log entries to CSV")
    p.add_argument('--output', '-o', required=True, help='Output CSV file path')
    p.add_argument('--action', help='Filter by action (INSERT, UPDATE, DELETE)')
    p.add_argument('--field', dest='field_name', help='Filter by field name')
    p.add_argument('--changed-by', help='Filter by user who made change')
    p.add_argument('--asset-id', type=int, help='Filter by asset id')
    p.add_argument('--asset-no', help='Filter by asset number')
    p.add_argument('--date-from', help='Filter start date (ISO format)')
    p.add_argument('--date-to', help='Filter end date (ISO format)')
    p.add_argument('--limit', type=int, help='Maximum number of rows to export')
    return p.parse_args()


def build_filters(args) -> Dict[str, Any]:
    filters: Dict[str, Any] = {}
    if args.action:
        filters['action'] = args.action
    if args.field_name:
        filters['field_name'] = args.field_name
    if args.changed_by:
        filters['changed_by'] = args.changed_by
    if args.asset_id is not None:
        filters['asset_id'] = args.asset_id
    if args.asset_no:
        filters['asset_no'] = args.asset_no
    if args.date_from:
        filters['date_from'] = args.date_from
    if args.date_to:
        filters['date_to'] = args.date_to
    return filters


def main():
    args = parse_args()
    filters = build_filters(args)

    db = database_service.get_database_instance()

    try:
        # Use underlying DB method to support limit
        count = db.export_audit_log(args.output, filters, args.limit)
        print(f"Wrote {count} audit entries to {args.output}")
    except Exception as e:
        print(f"Error exporting audit log: {e}")


if __name__ == '__main__':
    main()
