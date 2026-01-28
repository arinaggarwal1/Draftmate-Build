# engine/resolver.py

import re
from typing import Callable, List, Dict, Tuple


class PlaceholderResolver:
    """
    Resolves placeholders in email templates using CSV row data.

    Resolution priority:
    1. Derived values like first name, last name, firm, school
    2. Exact CSV header matches
    3. Empty string if unresolved

    Supported placeholders:
    {first name}
    {last name}
    {full name}
    {firm}
    {firm name}
    {school}
    Plus any exact CSV header name
    """

    COMPANY_OR_FIRM_SUBSTRS = ("company", "firm")
    SCHOOL_SUBSTRS = ("school", "college", "university", "uni")
    EMAIL_SUBSTR = "email"

    def __init__(
        self,
        headers_lower_ordered: List[str],
        row_dict_lower: Dict[str, str],
        parse_name_fn: Callable[[str], Tuple[str, str]],
    ):
        self.headers = headers_lower_ordered or []
        self.row = row_dict_lower or {}
        self._parse_name = parse_name_fn

        self._full_name_header = self._first_exact_header(("full name", "name"))
        # "Exact" match candidates for firm/company
        self._company_header = self._first_exact_header(("firm", "company", "firm name", "company name", "business"))
        if not self._company_header:
            # Fallback to substring but avoid "email" to prevent "Firm Email" being picked as firm
            self._company_header = self._first_contains_header(self.COMPANY_OR_FIRM_SUBSTRS, exclude_substrings=("email",))

        self._school_header = self._first_contains_header(self.SCHOOL_SUBSTRS)
        self._email_header = self._first_contains_header((self.EMAIL_SUBSTR,))

        self._derived = self._build_derived_defaults()

    # -------------------------
    # Header detection helpers
    # -------------------------

    def _first_exact_header(self, candidates):
        for h in self.headers:
            if h in candidates:
                return h
        return None

    def _first_contains_header(self, substrings, exclude_substrings=()):
        for h in self.headers:
            hl = h.lower()
            if any(sub in hl for sub in substrings):
                if exclude_substrings and any(ex in hl for ex in exclude_substrings):
                    continue
                return h
        return None

    # -------------------------
    # Derived values
    # -------------------------

    def _build_derived_defaults(self) -> Dict[str, str]:
        full_name_val = ""
        if self._full_name_header:
            full_name_val = self.row.get(self._full_name_header, "").strip()

        first, last = self._parse_name(full_name_val)

        firm_val = (
            self.row.get(self._company_header, "").strip()
            if self._company_header
            else ""
        )

        school_val = (
            self.row.get(self._school_header, "").strip()
            if self._school_header
            else ""
        )

        if not full_name_val and "name" in self.headers:
            fallback = self.row.get("name", "").strip()
            if fallback:
                full_name_val = fallback
                if not first and not last:
                    first, last = self._parse_name(fallback)

        derived = {
            "first name": first,
            "last name": last,
            "full name": full_name_val,
            "firm": firm_val,
            "firm name": firm_val,
            "school": school_val,
        }

        prefix_header = None
        for h in self.headers:
            if h == "prefix":
                prefix_header = h
                break

        if prefix_header:
            prefix_val = self.row.get(prefix_header, "").strip()
            if prefix_val:
                if last:
                    derived["first name"] = f"{prefix_val}. {last}"
                else:
                    derived["first name"] = f"{prefix_val}."
            else:
                if first:
                    derived["first name"] = first
        else:
            if first:
                derived["first name"] = first

        return derived

    # -------------------------
    # Placeholder resolution
    # -------------------------

    def _resolve_token(self, token_text: str) -> str:
        key = token_text.strip().lower()

        if key in self._derived:
            return self._derived[key]

        for h in self.headers:
            if h == key:
                return self.row.get(h, "").strip()

        return ""

    def resolve_text(self, text: str) -> str:
        if not text or "{" not in text:
            return text or ""

        pattern = re.compile(r"\{([^{}]+)\}", flags=re.IGNORECASE)

        def _sub(match):
            token = match.group(1)
            return self._resolve_token(token)

        return pattern.sub(_sub, text)

    # -------------------------
    # Convenience getters
    # -------------------------

    def get_email(self) -> str:
        if self._email_header:
            return self.row.get(self._email_header, "").strip()
        return self.row.get("email", "").strip()

    def get_firm(self) -> str:
        return self._derived.get("firm", "")

    def get_first_last(self) -> Tuple[str, str]:
        return (
            self._derived.get("first name", ""),
            self._derived.get("last name", ""),
        )

    def get_full_name(self) -> str:
        return self._derived.get("full name", "")
