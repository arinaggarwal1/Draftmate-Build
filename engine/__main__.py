#!/usr/bin/env python3
"""
DraftMate Engine CLI

Command-line interface for the DraftMate engine.
This serves as the bridge between Tauri frontend and Python engine.

Usage:
    python -m engine load-csv <path>
    python -m engine load-sheet <url>
    python -m engine preview --data <json> --templates <json> --overrides <json>
    python -m engine generate --data <json> --templates <json> --overrides <json> --subject <str> --resume <path>

All commands output JSON to stdout.
"""

import argparse
import json
import sys
import re
from typing import Any


def _parse_name(full_name: str) -> tuple[str, str]:
    """Parse full name into (first, last). Matches UI behavior."""
    if not full_name:
        return "", ""
    full_name = full_name.strip()
    if "," in full_name:
        last, first = [x.strip() for x in full_name.split(",", 1)]
        return first, last
    parts = full_name.split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[-1]


def _parse_bool(val: str) -> bool:
    """Parse string as boolean. Matches UI behavior."""
    return str(val).strip().lower() in {"1", "true", "yes", "y"}


def _is_generate_true(row: dict) -> bool:
    """Check if row should be generated."""
    # Create case-insensitive map
    lower_keys = {k.lower(): k for k in row.keys()}
    
    for key in ("generate", "gen"):
        if key in lower_keys:
            real_key = lower_keys[key]
            return _parse_bool(row[real_key])
    return True


def _is_email_valid(email: str) -> bool:
    """Basic email validation."""
    if not email or not isinstance(email, str):
        return False
    email = email.strip()
    if not email or "@" not in email:
        return False
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email))


def output_json(data: Any, success: bool = True) -> None:
    """Output JSON response to stdout."""
    response = {
        "success": success,
        "data": data if success else None,
        "error": None if success else data,
    }
    print(json.dumps(response, indent=2, ensure_ascii=False))


def cmd_load_csv(args: argparse.Namespace) -> None:
    """Load CSV file and return rows + headers."""
    from engine.data_sources import load_csv

    try:
        rows, headers = load_csv(args.path)
        output_json({"rows": rows, "headers": headers, "count": len(rows)})
    except Exception as e:
        output_json(str(e), success=False)


def cmd_load_sheet(args: argparse.Namespace) -> None:
    """Load Google Sheet and return rows + headers."""
    from engine.data_sources import load_google_sheet

    try:
        rows, headers = load_google_sheet(args.url)
        output_json({"rows": rows, "headers": headers, "count": len(rows)})
    except Exception as e:
        output_json(str(e), success=False)


def cmd_preview(args: argparse.Namespace) -> None:
    """Build preview rows from input data."""
    from engine.preview import build_preview_rows

    try:
        data = json.loads(args.data)
        rows = data.get("rows", [])
        headers = data.get("headers", [])
        templates = json.loads(args.templates)
        overrides = json.loads(args.overrides) if args.overrides else {}

        preview_rows = build_preview_rows(
            rows=rows,
            headers_lower=headers,
            templates=templates,
            recipient_template_overrides=overrides,
            parse_name_fn=_parse_name,
            only_recipients=args.only_recipients,
            is_generate_true_fn=_is_generate_true,
            is_email_valid_fn=_is_email_valid,
        )

        output_json({"preview_rows": preview_rows, "count": len(preview_rows)})
    except Exception as e:
        output_json(str(e), success=False)


def cmd_generate(args: argparse.Namespace) -> None:
    """Generate Outlook email drafts."""
    from engine.generator import generate_emails

    try:
        data = json.loads(args.data)
        rows = data.get("rows", [])
        headers = data.get("headers", [])
        templates = json.loads(args.templates)
        overrides = json.loads(args.overrides) if args.overrides else {}

        count = generate_emails(
            rows=rows,
            headers_lower=headers,
            templates=templates,
            recipient_template_overrides=overrides,
            parse_name_fn=_parse_name,
            is_generate_true_fn=_is_generate_true,
            is_email_valid_fn=_is_email_valid,
            subject_template=args.subject or "",
            resume_path=args.resume if args.resume else None,
            dry_run=args.dry_run,
        )

        output_json({"created": count})
    except Exception as e:
        output_json(str(e), success=False)


def cmd_read_files(args: argparse.Namespace) -> None:
    """Read multiple text files and return their contents."""
    from pathlib import Path

    try:
        files = []
        for filepath in args.paths:
            path = Path(filepath)
            if path.exists() and path.is_file():
                content = path.read_text(encoding="utf-8").rstrip()
                files.append({
                    "name": path.stem,  # filename without extension
                    "content": content,
                })
        output_json({"files": files})
    except Exception as e:
        output_json(str(e), success=False)


def cmd_export_templates(args: argparse.Namespace) -> None:
    """Export templates to a ZIP file in Downloads folder."""
    from pathlib import Path
    import zipfile

    try:
        templates = json.loads(args.templates)

        if not templates:
            output_json("No templates to export", success=False)
            return

        # Find Downloads folder
        downloads = Path.home() / "Downloads"
        dest_dir = downloads if downloads.exists() and downloads.is_dir() else Path.cwd()

        # Generate unique filename
        base_name = "My Email Templates"
        zip_path = dest_dir / f"{base_name}.zip"
        counter = 1
        while zip_path.exists():
            zip_path = dest_dir / f"{base_name} ({counter}).zip"
            counter += 1

        # Create safe filename function
        def safe_filename(name: str, max_len: int = 60) -> str:
            """Convert template name to safe filename."""
            import unicodedata
            s = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii")
            s = re.sub(r"[^\w\s-]", "", s)
            s = re.sub(r"[\s_-]+", "_", s).strip("_")
            if not s:
                s = "template"
            if len(s) > max_len:
                s = s[:max_len]
            return s

        # Create ZIP file
        with zipfile.ZipFile(zip_path, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for idx, tpl in enumerate(templates):
                text = tpl.get("text", "")
                name = tpl.get("name", f"Template {idx + 1}")
                fname = safe_filename(name)
                arcname = f"{idx + 1:02d}_{fname}.txt"
                zf.writestr(arcname, text)

        output_json({"path": zip_path.name})
    except Exception as e:
        output_json(str(e), success=False)


def cmd_validate_license(args: argparse.Namespace) -> None:
    """Validate a license key against the license database."""
    import urllib.request
    import ssl
    import csv
    import io
    import hashlib
    import uuid

    try:
        license_key = args.license_key
        if not license_key:
            output_json({"valid": False, "message": "No license key provided"})
            return

        # Get machine ID
        def get_machine_id():
            try:
                import subprocess
                result = subprocess.run(
                    ["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"],
                    capture_output=True, text=True
                )
                for line in result.stdout.split("\n"):
                    if "IOPlatformUUID" in line:
                        return hashlib.sha256(line.split('"')[-2].encode()).hexdigest()[:16]
            except Exception:
                pass
            return hashlib.sha256(str(uuid.getnode()).encode()).hexdigest()[:16]

        machine_id = get_machine_id()

        # Fetch license database
        SHEET_ID = "1-t4tZ4_AsQdP0LPidJaFKOmZrFGMz3lzl7_W7qGH_Iw"
        LICENSE_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"

        try:
            import certifi
            ctx = ssl.create_default_context(cafile=certifi.where())
        except ImportError:
            ctx = ssl.create_default_context()

        req = urllib.request.Request(LICENSE_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=12, context=ctx) as resp:
            data = resp.read().decode("utf-8-sig", errors="replace")

        rows = list(csv.DictReader(io.StringIO(data)))

        # Find matching license
        def normalize(row):
            return {
                (k or "").strip().lower().replace(" ", "_"): (v or "").strip()
                for k, v in (row.items() if row else [])
            }

        matched_row = None
        for r in rows:
            d = normalize(r)
            if d.get("license_key") == license_key:
                matched_row = d
                break

        if not matched_row:
            output_json({"valid": False, "message": "License key not found"})
            return

        # Check status
        status = (matched_row.get("status") or "").strip().lower()
        if status not in {"active", "1", "true", "yes", "enabled"}:
            output_json({"valid": False, "message": f"License is not active (status: {status})"})
            return

        # Check machine binding
        sheet_mid = (matched_row.get("machine_id") or "").strip()
        if not sheet_mid:
            # Not bound yet - try to bind
            BIND_URL = "https://script.google.com/macros/s/AKfycbwJ3t0--FrAW6ANHMryUY5tDk1u3FzWWK0Q2JbsbY7JR_ulv9iHl9T_WD5Iw6LRAuvI/exec"
            SHARED_SECRET = "5c1a8a9f-44b2-4f0a-9b10-1f93b90f7a73"

            payload = json.dumps({
                "action": "bind",
                "license_key": license_key,
                "machine_id": machine_id,
                "secret": SHARED_SECRET,
            }).encode("utf-8")

            bind_req = urllib.request.Request(
                BIND_URL,
                data=payload,
                headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"},
                method="POST",
            )
            try:
                with urllib.request.urlopen(bind_req, timeout=12, context=ctx) as resp:
                    bind_resp = json.loads(resp.read().decode("utf-8", errors="replace"))
                    if bind_resp.get("ok"):
                        output_json({"valid": True, "message": "License validated and bound to this device"})
                    else:
                        output_json({"valid": False, "message": bind_resp.get("msg", "Binding failed")})
            except Exception as e:
                output_json({"valid": False, "message": f"Failed to bind license: {e}"})
            return

        if sheet_mid == machine_id:
            output_json({"valid": True, "message": "License validated for this device"})
        else:
            short = f"{sheet_mid[:6]}..." if len(sheet_mid) > 6 else sheet_mid
            output_json({"valid": False, "message": f"License is bound to another device ({short})"})

    except Exception as e:
        output_json({"valid": False, "message": f"License validation failed: {e}"})


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="draftmate",
        description="DraftMate Engine CLI - JSON bridge for Tauri frontend",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # load-csv
    p_csv = subparsers.add_parser("load-csv", help="Load data from CSV file")
    p_csv.add_argument("path", help="Path to CSV file")
    p_csv.set_defaults(func=cmd_load_csv)

    # load-sheet
    p_sheet = subparsers.add_parser("load-sheet", help="Load data from Google Sheet")
    p_sheet.add_argument("url", help="Google Sheets URL")
    p_sheet.set_defaults(func=cmd_load_sheet)

    # preview
    p_preview = subparsers.add_parser("preview", help="Build preview rows")
    p_preview.add_argument("--data", required=True, help="JSON with rows and headers")
    p_preview.add_argument("--templates", required=True, help="JSON array of templates")
    p_preview.add_argument("--overrides", default="{}", help="JSON object of email->template_id overrides")
    p_preview.add_argument("--only-recipients", action="store_true", default=True, help="Filter to recipients only")
    p_preview.add_argument("--all-rows", dest="only_recipients", action="store_false", help="Include all rows with emails")
    p_preview.set_defaults(func=cmd_preview)

    # generate
    p_gen = subparsers.add_parser("generate", help="Generate Outlook drafts")
    p_gen.add_argument("--data", required=True, help="JSON with rows and headers")
    p_gen.add_argument("--templates", required=True, help="JSON array of templates")
    p_gen.add_argument("--overrides", default="{}", help="JSON object of email->template_id overrides")
    p_gen.add_argument("--subject", required=True, help="Subject line template")
    p_gen.add_argument("--resume", help="Path to resume PDF")
    p_gen.add_argument("--dry-run", action="store_true", help="Don't create drafts, just count")
    p_gen.set_defaults(func=cmd_generate)

    # read-files
    p_read = subparsers.add_parser("read-files", help="Read multiple text files")
    p_read.add_argument("paths", nargs="+", help="Paths to text files")
    p_read.set_defaults(func=cmd_read_files)

    # export-templates
    p_export = subparsers.add_parser("export-templates", help="Export templates to ZIP")
    p_export.add_argument("--templates", required=True, help="JSON array of templates")
    p_export.set_defaults(func=cmd_export_templates)

    # validate-license
    p_license = subparsers.add_parser("validate-license", help="Validate a license key")
    p_license.add_argument("license_key", help="License key to validate")
    p_license.set_defaults(func=cmd_validate_license)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
