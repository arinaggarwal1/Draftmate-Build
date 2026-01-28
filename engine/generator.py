# engine/generator.py

import subprocess
from typing import List, Dict, Callable

from engine.resolver import PlaceholderResolver
from engine.preview import build_preview_rows


# ============================================================
# helpers
# ============================================================

def _escape_applescript(s: str) -> str:
    if not s:
        return ""
    return s.replace("\\", "\\\\").replace('"', '\\"')


def _wrap_in_html(text: str) -> str:
    """Wrap plain text in HTML formatting for email body."""
    text = text.strip()
    html_content = text.replace("  ", "<br><br>").replace("\n", "<br>").strip()
    html_content = html_content.lstrip("<br>").rstrip("<br>")
    return f"<html><body style='margin:0;padding:0;font-family:Arial,sans-serif;'>{html_content}</body></html>".strip()


def _run_applescript(script: str) -> None:
    subprocess.run(
        ["osascript", "-e", script],
        check=True,
    )


def _build_rows_by_email(
    rows: List[Dict],
    headers_lower: List[str],
    parse_name_fn: Callable[[str], tuple[str, str]],
) -> Dict[str, Dict]:
    """Build email -> row mapping internally."""
    result = {}
    for r in rows:
        resolver = PlaceholderResolver(headers_lower, r, parse_name_fn)
        email = (resolver.get_email() or "").lower().strip()
        if email:
            result[email] = r
    return result


# ============================================================
# public API
# ============================================================

def generate_emails(
    rows: List[Dict],
    headers_lower: List[str],
    templates: List[Dict],
    recipient_template_overrides: Dict[str, str],
    parse_name_fn: Callable[[str], tuple[str, str]],
    is_generate_true_fn: Callable[[Dict], bool],
    is_email_valid_fn: Callable[[str], bool],
    *,
    subject_template: str,
    resume_path: str | None,
    dry_run: bool = False,
) -> int:
    """
    Generates Outlook drafts.

    Builds preview rows and email mapping internally.
    UI only needs to pass raw data and callbacks.
    """

    # Build preview rows internally
    preview_rows = build_preview_rows(
        rows=rows,
        headers_lower=headers_lower,
        templates=templates,
        recipient_template_overrides=recipient_template_overrides,
        parse_name_fn=parse_name_fn,
        only_recipients=True,
        is_generate_true_fn=is_generate_true_fn,
        is_email_valid_fn=is_email_valid_fn,
    )

    # Build email -> row mapping internally
    rows_by_email = _build_rows_by_email(rows, headers_lower, parse_name_fn)

    count = 0
    templates_by_id = {t["id"]: t for t in templates}

    for p in preview_rows:
        # Use normalized (lowercase) email for lookup since rows_by_email uses lowercase keys
        email_display = p.get("email") or ""
        email_norm = p.get("email_norm") or email_display.lower().strip()
        tid = p.get("template_id")

        if not email_norm or not tid:
            continue

        row = rows_by_email.get(email_norm)
        if not row:
            continue

        tpl = templates_by_id.get(tid)
        if not tpl:
            continue

        resolver = PlaceholderResolver(headers_lower, row, parse_name_fn)

        subject = resolver.resolve_text(subject_template or "")
        body_plain = resolver.resolve_text(tpl.get("text", ""))
        body = _wrap_in_html(body_plain)

        if dry_run:
            count += 1
            continue

        _create_outlook_draft(
            to=email_display or email_norm,  # Use display email for Outlook
            subject=subject,
            body=body,
            resume_path=resume_path,
        )

        count += 1

    return count


# ============================================================
# Outlook integration
# ============================================================

def _create_outlook_draft(
    *,
    to: str,
    subject: str,
    body: str,
    resume_path: str | None,
) -> None:
    """
    Creates a draft email in Outlook (Classic) via AppleScript.
    """

    subj = _escape_applescript(subject)
    bod = _escape_applescript(body)
    to_esc = _escape_applescript(to)

    attach_cmd = ""
    if resume_path:
        esc_attach = _escape_applescript(resume_path)
        attach_cmd = f'make new attachment with properties {{file:POSIX file "{esc_attach}"}}'

    script = f'''
    on run
        try
            tell application "Microsoft Outlook" to get name
        on error errMsg number errNum
            error "Outlook AppleScript not available. If you're on 'New Outlook', switch to Classic Outlook. " & errMsg number errNum
        end try

        tell application "Microsoft Outlook"
            set newMessage to make new outgoing message
            tell newMessage
                make new recipient at end of to recipients with properties {{email address:{{address:"{to_esc}"}}}}
                set subject to "{subj}"
                set content to "{bod}"
                {attach_cmd}
                open
            end tell
            activate
        end tell
    end run
    '''

    _run_applescript(script)
