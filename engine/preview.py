# engine/preview.py

from typing import List, Dict, Tuple, Callable, Optional

from engine.resolver import PlaceholderResolver


# ============================================================
# helpers
# ============================================================

def _template_by_id(templates: list[dict], tid: str) -> Optional[dict]:
    for t in templates:
        if t.get("id") == tid:
            return t
    return None


def _rotatable_templates(templates: list[dict]) -> list[dict]:
    return [t for t in templates if not t.get("manual_only", False)]


# ============================================================
# core template selection logic
# ============================================================

def choose_template_for_row(
    row: dict,
    headers_lower: list[str],
    templates: list[dict],
    recipient_template_overrides: dict[str, str],
    firm_counts: dict[str, int],
    parse_name_fn: Callable[[str], tuple[str, str]],
) -> Tuple[Optional[dict], bool]:
    """
    Returns:
        (template_dict or None, is_manual_override)
    """

    resolver = PlaceholderResolver(headers_lower, row, parse_name_fn)

    email = (resolver.get_email() or "").lower().strip()
    firm = resolver.get_firm() or ""

    # ------------------------
    # manual override first
    # ------------------------
    if email and email in recipient_template_overrides:
        tid = recipient_template_overrides.get(email)
        t = _template_by_id(templates, tid)
        if t:
            return t, True

    # ------------------------
    # automatic rotation
    # ------------------------
    rot = _rotatable_templates(templates)
    if not rot:
        return None, False

    firm_key = firm or ""
    firm_counts[firm_key] = firm_counts.get(firm_key, 0) + 1

    idx = (firm_counts[firm_key] - 1) % len(rot)
    return rot[idx], False


# ============================================================
# preview row construction
# ============================================================

def build_preview_rows(
    rows: list[dict],
    headers_lower: list[str],
    templates: list[dict],
    recipient_template_overrides: dict[str, str],
    parse_name_fn: Callable[[str], tuple[str, str]],
    *,
    only_recipients: bool = True,
    is_generate_true_fn: Callable[[dict], bool],
    is_email_valid_fn: Callable[[str], bool],
) -> list[dict]:
    """
    Build preview table rows.

    Returns list of dicts with keys:
        name
        email            (display value, original casing)
        email_norm       (lower/strip, stable identity)
        firm
        template_name
        template_id
        is_manual
    """

    out: list[dict] = []
    firm_counts: dict[str, int] = {}

    for row in rows:
        resolver = PlaceholderResolver(headers_lower, row, parse_name_fn)

        email_display = resolver.get_email() or ""
        email_norm = email_display.lower().strip()

        # Behavior must match UI:
        # - If only_recipients is ON: show only eligible rows
        # - If only_recipients is OFF: show all provided rows, but only assign templates
        #   when generate==true AND email is valid
        eligible = bool(is_generate_true_fn(row) and is_email_valid_fn(email_display))

        if only_recipients and not eligible:
            continue

        firm = resolver.get_firm() or ""
        first, last = resolver.get_first_last()
        full_name = resolver.get_full_name()

        prefix_val = row.get("prefix", "").strip()
        if prefix_val:
            name_display = f"{first} ({full_name})"
        else:
            name_display = (first + (" " + last if last else "")).strip()

        if eligible:
            chosen, is_manual = choose_template_for_row(
                row=row,
                headers_lower=headers_lower,
                templates=templates,
                recipient_template_overrides=recipient_template_overrides,
                firm_counts=firm_counts,
                parse_name_fn=parse_name_fn,
            )
        else:
            chosen, is_manual = None, False

        if chosen:
            tpl_name = chosen.get("name", "")
            tpl_id = chosen.get("id")
        else:
            tpl_name = "–"
            tpl_id = None
            is_manual = False

        out.append(
            {
                "name": name_display,
                "email": email_display,
                "email_norm": email_norm,
                "firm": firm,
                "template_name": tpl_name,
                "template_id": tpl_id,
                "is_manual": bool(is_manual),
                "is_eligible": eligible,
            }
        )

    return out
