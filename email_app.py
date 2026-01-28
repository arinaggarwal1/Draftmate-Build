#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DraftMate — CustomTkinter UI

This application helps generate personalized email drafts for Investment Banking networking outreach.

Key Features:
- Multiple profiles: Create unlimited profiles for different outreach strategies
- Dual data sources: Google Sheets (public URLs) or local CSV files
- Template system: Create reusable email templates with placeholder support
- Resume integration: Upload PDF resume for context in emails
- Outlook integration: Generate drafts directly in Microsoft Outlook
- License management: Secure licensing system with device binding
- Export functionality: Export templates as ZIP files to Downloads folder

Placeholder System:
- Basic placeholders: {first name}, {last name}, {full name}, {firm}, {school}
- CSV column matching: Any exact CSV header can be used as a placeholder
- Special handling: Prefix column modifies {first name} display

Requirements: pip install customtkinter
"""

from engine.data_sources import load_csv, load_google_sheet
from engine.preview import build_preview_rows
from engine.generator import generate_emails
from engine.resolver import PlaceholderResolver

import os
import csv
import io
import json
import re
import subprocess
import sys
import urllib.parse
import urllib.request
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, simpledialog, ttk, messagebox
import tkinter.font as tkfont
import zipfile
import webbrowser
import uuid


import tempfile
from pathlib import Path

import ssl
import certifi

import customtkinter as ctk


import hashlib
import uuid

import markdown
from bs4 import BeautifulSoup

import subprocess, pathlib, re, platform
from appdirs import user_data_dir

APP_NAME   = "IB Email Generator"
APP_AUTHOR = "Arin Aggarwal"
APP_DIR    = user_data_dir(APP_NAME, APP_AUTHOR)
os.makedirs(APP_DIR, exist_ok=True)
INSTALL_ID_FILE = os.path.join(APP_DIR, ".install_id")

# Development flag to skip license validation (set environment variable DEV_SKIP_LICENSE=1)
DEV_SKIP_LICENSE = os.environ.get("DEV_SKIP_LICENSE") == "1"

# Application metadata for settings storage
APP_NAME = "IB Email Generator"
APP_AUTHOR = "Arin Aggarwal"

# Settings file location using platform-appropriate user data directory
SETTINGS_FILE = os.path.join(
    user_data_dir(APP_NAME, APP_AUTHOR),
    "email_app_settings.json"
)

# Ensure settings directory exists
os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)

# Default subject template with placeholder for firm name
DEFAULT_SUBJECT = "Duke Student interested in IB at {firm}"

# Platform detection for macOS-specific features
IS_MAC = sys.platform == "darwin"

# Email validation regex pattern
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# Settings schema version for future compatibility
SCHEMA_VERSION = 1


# -------------------------
# Toast Notification System
# -------------------------
class ToastManager:
    """
    Manages toast notifications that appear in the bottom-right corner of the app.
    Supports different message types with color coding and automatic dismissal.
    """
    
    # Color schemes for different notification types
    COLORS = {
        "info":    {"bg": "#1f2937", "fg": "#ffffff"},  # Dark blue for general info
        "success": {"bg": "#065f46", "fg": "#ffffff"},  # Dark green for success
        "warning": {"bg": "#92400e", "fg": "#ffffff"},  # Dark orange for warnings
        "error":   {"bg": "#7f1d1d", "fg": "#ffffff"},  # Dark red for errors
    }

    def __init__(self, root: ctk.CTk):
        """Initialize toast manager with reference to main window."""
        self.root = root
        self.active = []  # Track currently displayed toasts
        self.recipient_template_overrides = {}

    def show(self, message: str, kind: str = "info", duration_ms: int = 2400):
        """
        Display a toast notification.
        
        Args:
            message: Text to display in the toast
            kind: Type of notification (info, success, warning, error)
            duration_ms: How long to show the toast before auto-dismissing
        """
        # Get color configuration for the notification type
        cfg = self.COLORS.get(kind, self.COLORS["info"])
        
        # Create a topmost window for the toast
        toast = ctk.CTkToplevel(self.root)
        toast.overrideredirect(True)  # Remove window decorations
        toast.lift()
        toast.attributes("-topmost", True)  # Keep on top of all windows
        
        # Set transparency if supported by the platform
        try:
            toast.wm_attributes("-alpha", 0.97)
        except Exception:
            pass

        # Create the toast content with styled frame and label
        frame = ctk.CTkFrame(toast, fg_color=cfg["bg"], corner_radius=10)
        frame.pack(fill="both", expand=True)
        label = ctk.CTkLabel(frame, text=message, text_color=cfg["fg"], wraplength=380, justify="left")
        label.pack(padx=14, pady=10)

        # Calculate toast size and position (bottom-right corner with stacking)
        self.root.update_idletasks()
        rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
        rw, rh = self.root.winfo_width(), self.root.winfo_height()
        
        # Dynamic width based on message length, with min/max constraints
        width = min(420, max(260, int(12 * (len(message) ** 0.55))))
        height = label.winfo_reqheight() + 24
        
        # Stack toasts vertically with 10px spacing
        offset_y = 16 + (len(self.active) * (height + 10))
        x = rx + rw - width - 16
        y = ry + rh - height - offset_y
        toast.geometry(f"{width}x{height}+{x}+{y}")

        # Add to active toasts list
        self.active.append(toast)

        def _destroy():
            """Remove toast and reposition remaining toasts."""
            if toast in self.active:
                idx = self.active.index(toast)
                toast.destroy()
                self.active.pop(idx)
                self._reposition()

        # Schedule automatic dismissal
        self.root.after(duration_ms, _destroy)

    def _reposition(self):
        """Reposition all active toasts when one is removed."""
        self.root.update_idletasks()
        rx, ry = self.root.winfo_rootx(), self.root.winfo_rooty()
        rw, rh = self.root.winfo_width(), self.root.winfo_height()
        cur_y = 16
        
        # Update position of each remaining toast
        for t in list(self.active):
            try:
                w = t.winfo_width() or 320
                h = t.winfo_height() or 40
                x = rx + rw - w - 16
                y = ry + rh - h - cur_y
                t.geometry(f"+{x}+{y}")
                cur_y += h + 10
            except Exception:
                pass  # Handle destroyed windows gracefully

# -------------------------
# License Management System
# -------------------------

def _sh(cmd):
    try:
        out = subprocess.check_output(cmd, stderr=subprocess.DEVNULL)
        return out.decode("utf-8", errors="ignore")
    except Exception:
        return ""

def _mac_io_platform_uuid():
    txt = _sh(["ioreg", "-rd1", "-c", "IOPlatformExpertDevice"])
    m = re.search(r'"IOPlatformUUID"\s=\s"([0-9A-Fa-f-]+)"', txt)
    return (m.group(1) if m else "").strip()

def _mac_serial_number():
    txt = _sh(["system_profiler", "SPHardwareDataType"])
    m = re.search(r"Serial Number \(system\):\s*([A-Za-z0-9]+)", txt)
    return (m.group(1) if m else "").strip()

def _ensure_install_id():
    """
    Store a random, per-installation ID.
    Tries Keychain first (via keyring), falls back to file.
    """
    install_id = ""
    try:
        import keyring
        service = f"{APP_NAME}-install-id"
        account = "default"
        install_id = keyring.get_password(service, account) or ""
        if not install_id:
            install_id = str(uuid.uuid4())
            keyring.set_password(service, account, install_id)
    except Exception:
        try:
            if not os.path.exists(INSTALL_ID_FILE):
                pathlib.Path(INSTALL_ID_FILE).write_text(str(uuid.uuid4()))
            install_id = pathlib.Path(INSTALL_ID_FILE).read_text().strip()
        except Exception:
            install_id = str(uuid.uuid4())
    return install_id

def get_machine_id() -> str:
    """
    macOS-only device fingerprint.
    Combines Hardware UUID + Serial + persisted Installation ID.
    """
    assert platform.system().lower() == "darwin", "This fingerprint is for macOS"
    parts = [
        _mac_io_platform_uuid(),
        _mac_serial_number(),
        _ensure_install_id()
    ]
    norm = "|".join(p.lower() for p in parts if p)
    if not norm:
        norm = str(uuid.getnode())  # fallback
    digest = hashlib.sha256(norm.encode("utf-8")).hexdigest()
    return digest[:16]


class LicenseManager:
    """
    Manages license validation and device binding.
    
    Workflow:
    1. Validates license key against Google Sheets database
    2. Checks if license is active and not bound to another device
    3. Binds license to current device if unbound
    4. Maintains device binding for subsequent validations
    """

    # Google Sheets configuration for license database
    SHEET_ID = "1-t4tZ4_AsQdP0LPidJaFKOmZrFGMz3lzl7_W7qGH_Iw"
    LICENSE_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid=0"

    # Apps Script endpoint for writing device bindings back to the sheet
    BIND_URL = "https://script.google.com/macros/s/AKfycbwJ3t0--FrAW6ANHMryUY5tDk1u3FzWWK0Q2JbsbY7JR_ulv9iHl9T_WD5Iw6LRAuvI/exec"

    # Shared secret for authenticating with Apps Script (acts like a password)
    SHARED_SECRET = "5c1a8a9f-44b2-4f0a-9b10-1f93b90f7a73"

    def __init__(self, app):
        """Initialize license manager with app reference and current device ID."""
        self.app = app
        self.machine_id = get_machine_id()

    def _fetch_rows(self):
        """
        Download license database from Google Sheets as CSV.
        Returns list of dictionaries representing each row.
        """
        req = urllib.request.Request(
            self.LICENSE_SHEET_URL,
            headers={"User-Agent": "Mozilla/5.0"}  # Required for Google Sheets access
        )
        # Use proper SSL context with certificate verification
        ctx = ssl.create_default_context(cafile=certifi.where())
        with urllib.request.urlopen(req, timeout=12, context=ctx) as resp:
            data = resp.read().decode("utf-8-sig", errors="replace")  # Handle BOM
        return list(csv.DictReader(io.StringIO(data)))

    def _normalize_headers(self, row_dict):
        """
        Normalize dictionary keys to lowercase with underscores.
        Trims whitespace from both keys and values.
        """
        return {
            (k or "").strip().lower().replace(" ", "_"): (v or "").strip()
            for k, v in (row_dict.items() if row_dict else [])
        }

    def _find_row_for_key(self, rows, license_key):
        """
        Find the row matching the given license key (case-sensitive).
        Returns normalized dictionary or None if not found.
        """
        for r in rows:
            d = self._normalize_headers(r)
            if d.get("license_key") == license_key:  # Case-sensitive exact match
                return d
        return None

    def check_license_verbose(self, license_key: str) -> tuple[bool, str, dict]:
        """
        Comprehensive license validation check.
        
        Returns:
            tuple: (is_valid, status_message, matched_row_data)
            - is_valid: True only if license is active and available for this device
            - status_message: Human-readable explanation of the result
            - matched_row_data: Raw data from the license database
        """
        if not license_key:
            return False, "No license key entered.", {}

        try:
            rows = self._fetch_rows()
        except Exception as e:
            return False, f"License check failed to reach the sheet: {e}", {}

        # Find matching license entry
        row = self._find_row_for_key(rows, license_key)
        if not row:
            return False, "No matching license entry found for this key.", {}

        # Check if license is active
        status = (row.get("status") or "").strip().lower()
        if status not in {"active", "1", "true", "yes", "enabled"}:
            return False, f"License found but status='{row.get('status','')}'.", row

        # Check device binding status
        sheet_mid = (row.get("machine_id") or "").strip()
        my_mid = self.machine_id

        if not sheet_mid:
            # License exists and is active, but not bound to any device yet
            return True, "License valid (not yet bound to any device).", row

        if sheet_mid == my_mid:
            # License is bound to this specific device
            return True, "License valid and already bound to this device.", row

        # License is bound to a different device
        short = f"{sheet_mid[:6]}…" if len(sheet_mid) > 6 else sheet_mid
        return False, f"License is bound to another device ({short}).", row

    def bind_machine_id(self, license_key: str) -> tuple[bool, str]:
        """
        Bind the current device to the specified license key.
        
        Makes a POST request to Apps Script which:
        1. Verifies the shared secret
        2. Locates the license key row
        3. Updates machine_id if empty or matches current device
        
        Returns:
            tuple: (success, status_message)
        """
        payload = {
            "action": "bind",
            "license_key": license_key,
            "machine_id": self.machine_id,
            "secret": self.SHARED_SECRET,
        }
        
        data = json.dumps(payload).encode("utf-8")
        req = urllib.request.Request(
            self.BIND_URL,
            data=data,
            headers={"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"},
            method="POST",
        )
        
        ctx = ssl.create_default_context(cafile=certifi.where())
        try:
            with urllib.request.urlopen(req, timeout=12, context=ctx) as resp:
                body = resp.read().decode("utf-8", errors="replace")
            
            # Parse JSON response from Apps Script
            try:
                j = json.loads(body)
            except Exception:
                return False, f"Bind: invalid JSON response: {body[:160]}"
            
            ok = bool(j.get("ok"))
            msg = str(j.get("msg") or ("Bind ok" if ok else "Bind failed"))
            return ok, msg
        except Exception as e:
            return False, f"Bind request failed: {e}"

    def validate_and_bind(self, license_key: str) -> tuple[bool, str]:
        """
        Complete license validation and binding workflow.
        
        Process:
        1. Check current license status
        2. If valid but unbound, attempt to bind to this device
        3. Re-verify after binding to confirm success
        
        Returns:
            tuple: (final_success, final_message)
        """
        # Initial validation check
        ok, why, row = self.check_license_verbose(license_key)
        if not ok:
            return False, why

        # Check if binding is needed
        sheet_mid = (row.get("machine_id") or "").strip()
        if not sheet_mid:
            # License is valid but unbound - attempt to bind
            b_ok, b_msg = self.bind_machine_id(license_key)
            if not b_ok:
                return False, f"License valid but binding failed: {b_msg}"
            
            # Verify binding was successful
            ok2, why2, _ = self.check_license_verbose(license_key)
            if ok2:
                return True, "License validated and bound to this device."
            return False, f"License bound but re-check failed: {why2}"

        # License was already bound to this device
        return True, "License validated for this device."


# ============================================
# Placeholder Resolution System
# ============================================

# ============================================
# Main Application Class
# ============================================
class EmailApp(ctk.CTk):
    """
    Main application class for DraftMate.
    
    Features:
    - Multi-profile support for different outreach strategies
    - Google Sheets and CSV data source integration
    - Template management with placeholder support
    - License validation and device binding
    - Outlook integration for draft creation
    - Resume PDF attachment handling
    """

    def _show_machine_id(self):
        """Display and copy machine ID to clipboard for license support."""
        mid = get_machine_id()
        self.clipboard_clear()
        self.clipboard_append(mid)
        messagebox.showinfo("Machine ID", f"Machine ID (copied):\n\n{mid}")

    def __init__(self):
        """Initialize the main application window and all components."""
        super().__init__()
        
        # Configure CustomTkinter appearance
        ctk.set_appearance_mode("system")  # Follow system dark/light mode
        ctk.set_default_color_theme("blue")
        
        # Main window setup
        self.title("DraftMate")
        self.geometry("1240x780")
        self.resizable(False, False)

        # Initialize toast notification system
        self.toast = ToastManager(self)

        # ========== APPLICATION STATE ==========
        
        # Profile management
        self.profiles = {}  # Dictionary of profile_name -> profile_data
        self.profile_order = []  # Ordered list of profile names
        self.active_profile = tk.StringVar(value="Default")

        # Per-profile data binding (these track current profile's state)
        self.csv_path = tk.StringVar()  # Path to local CSV file
        self.resume_path = tk.StringVar()  # Path to resume PDF
        self.subject_template = tk.StringVar(value=DEFAULT_SUBJECT)  # Email subject template
        self.data_source = tk.StringVar(value="sheet")  # "sheet" or "csv"
        self.gsheet_url = tk.StringVar()  # Google Sheets URL
        
        # Global license key (shared across all profiles)
        self.license_key = tk.StringVar()

        # Data caches for performance
        self._gsheet_rows_cache = None  # Cached Google Sheets data
        self._last_headers_lower = []  # Cached column headers

        # Template management
        self.templates = []  # List of template dictionaries
        self.current_index = None  # Currently selected template index
        self.unsaved_buffers = {}  # Unsaved changes by template index
        self.recipient_template_overrides = {}


        # UI configuration for template listbox
        self.listbox_char_width = len("EmailEmailEmailEmailEmailEmailEmail")
        self.listbox_font_family = "Menlo" if IS_MAC else "Courier New"
        self.listbox_font_size = 13
        self.listbox_height = 7

        # Load saved settings from disk (before UI creation)
        self._load_settings()

        # Build the complete user interface
        self._build_ui()

        # ========== LICENSE SYSTEM INITIALIZATION ==========
        self._licensed = False  # Track current license status
        self.license_mgr = LicenseManager(self)

        # Automatic license validation on startup
        self._check_license_smart(force_online=True)

        # Keyboard shortcuts
        self.bind_all("<Command-s>" if IS_MAC else "<Control-s>", self._kb_save)
        self.bind_all("<F2>", self._kb_rename)

        # Handle window closing
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Preview window state
        self._preview_win = None
        self._preview_only_recipients = tk.BooleanVar(value=True)
        self._preview_item_to_email = {}
        self._preview_cell_editor = None


    def _check_license_smart(self, force_online=False):
        """
        Smart license validation with local and online checking.
        
        Local validation happens every time (fast).
        Online validation only occurs:
        - On startup (force_online=True)
        - Every 20+ minutes during use
        
        This reduces API calls while maintaining security.
        """
        import time
        
        key = (self.license_key.get() or "").strip()
        if not key:
            self._licensed = False
            self._update_license_badge()
            self._apply_license_gate()
            return False

        # Always perform local validation (checks cached sheet data)
        ok, why, row = self.license_mgr.check_license_verbose(key)
        if not ok:
            self._licensed = False
            self.toast.show(f"License invalid: {why}", kind="error")
            self._update_license_badge()
            self._apply_license_gate()
            return False

        # Online validation with time-based throttling
        now = time.time()
        time_threshold = 0.01 * 60  # 0.01 minutes = 0.6 seconds (for testing)
        
        if force_online or (now - getattr(self, '_last_online_check', 0) > time_threshold):
            ok2, msg = self.license_mgr.validate_and_bind(key)
            self._licensed = bool(ok2)
            self._last_online_check = now
            self.toast.show(msg, kind="success" if ok2 else "error")

        self._update_license_badge()
        self._apply_license_gate()
        return self._licensed

    # =========================
    # User Interface Construction
    # =========================
    def _build_ui(self):
        """
        Build the complete user interface layout.
        
        Layout structure:
        - Row 0: Profile management top bar
        - Row 1: Separator line
        - Row 2: Main content area (data sources, templates, editor)
        """
        # Configure main grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # --------- TOP BAR: Profile Management ----------
        topbar = ctk.CTkFrame(self, fg_color=("#0e1726", "#0e1726"))
        topbar.grid(row=0, column=0, sticky="ew")
        topbar.grid_columnconfigure(0, weight=1)

        # Inner container for profile controls
        inner = ctk.CTkFrame(topbar, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=(10, 10))

        # Profile label and dropdown
        prof_label = ctk.CTkLabel(inner, text="Profile", font=ctk.CTkFont(size=14, weight="bold"))
        prof_label.pack(side="left")

        self.profile_menu = ctk.CTkOptionMenu(
            inner,
            values=self.profile_order if self.profile_order else ["Default"],
            command=self._on_profile_selected_ui,
            width=200
        )
        self.profile_menu.pack(side="left", padx=(6, 10))
        self.profile_menu.set(self.active_profile.get())

        # Profile management buttons
        ctk.CTkButton(inner, text="New", width=70, command=self._profile_new).pack(side="left")
        ctk.CTkButton(inner, text="Duplicate", width=100, command=self._profile_duplicate).pack(side="left", padx=(6, 0))
        ctk.CTkButton(inner, text="Rename", width=90, command=self._profile_rename).pack(side="left", padx=(6, 0))
        ctk.CTkButton(inner, text="Delete", width=80, fg_color="#7f1d1d", hover_color="#991b1b",
                      command=self._profile_delete).pack(side="left", padx=(6, 0))
        
        # Help and license buttons
        ctk.CTkButton(
            inner,
            text="ℹ︎ Help",
            width=70,
            command=self._open_help_modal
        ).pack(side="left", padx=(6, 0))

        self.license_btn = ctk.CTkButton(inner, text="License…", width=100, command=self._open_license_modal)
        self.license_btn.pack(side="left", padx=(6, 0))

        # License status indicator
        self.license_status_top = ctk.CTkLabel(inner, text="● Unlicensed", text_color="#f87171")
        self.license_status_top.pack(side="left", padx=(8, 0))

        # Visual separator under top bar
        sep = ctk.CTkFrame(self, height=1, fg_color=("#1f2937", "#1f2937"))
        sep.grid(row=1, column=0, sticky="ew")

        # --------- MAIN CONTENT AREA ----------
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 12))
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        # Top section: Data source controls
        controls = ctk.CTkFrame(main, fg_color="transparent")
        controls.grid(row=0, column=0, sticky="ew", pady=(0,0))
        for i in range(12):
            controls.grid_columnconfigure(i, weight=1)

        # Data source selection (Google Sheet vs Local CSV)
        ctk.CTkLabel(controls, text="Data Source").grid(row=0, column=0, sticky="w", pady=(8,8))
        self.source_seg = ctk.CTkSegmentedButton(
            controls, values=["Google Sheet", "Local CSV"], command=self._on_change_data_source
        )
        self.source_seg.grid(row=0, column=1, columnspan=2, sticky="w", padx=(6, 12))
        self.source_seg.set("Google Sheet" if self.data_source.get() == "sheet" else "Local CSV")

        # Google Sheets controls row
        self.gsheet_label = ctk.CTkLabel(controls, text="Google Sheets URL (public)")
        self.gsheet_label.grid(row=1, column=0, sticky="w", pady=(0,0))
        self.gsheet_entry = ctk.CTkEntry(
            controls, textvariable=self.gsheet_url,
            placeholder_text="Paste an 'Anyone with the link – Viewer' URL…"
        )
        self.gsheet_entry.grid(row=1, column=1, columnspan=6, sticky="ew", padx=6)

        self.gsheet_load_btn = ctk.CTkButton(
            controls, text="Load/Refresh", command=lambda: self._load_gsheet(quiet=False)
        )
        self.gsheet_load_btn.grid(row=1, column=7, sticky="ew")

        self.gsheet_clear_btn = ctk.CTkButton(
            controls, text="Clear", command=self._clear_gsheet_url
        )
        self.gsheet_clear_btn.grid(row=1, column=8, sticky="ew", padx=(6, 0))

        self.gsheet_link_btn = ctk.CTkButton(
            controls, text="🔗", width=36, height=28, command=self._open_gsheet_in_browser
        )
        self.gsheet_link_btn.grid(row=1, column=9, sticky="w", padx=(6, 0))

        # CSV file controls row
        self.csv_label = ctk.CTkLabel(controls, text="CSV File")
        self.csv_label.grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.csv_entry = ctk.CTkEntry(controls, textvariable=self.csv_path, placeholder_text="Select a CSV…")
        self.csv_entry.grid(row=2, column=1, columnspan=6, sticky="ew", padx=6, pady=(8, 0))
        self.csv_browse_btn = ctk.CTkButton(controls, text="Browse", command=self._browse_csv)
        self.csv_browse_btn.grid(row=2, column=7, sticky="ew", pady=(8, 0))
        self.csv_clear_btn = ctk.CTkButton(controls, text="Clear", command=self._clear_csv_path)
        self.csv_clear_btn.grid(row=2, column=8, sticky="ew", padx=(6, 0), pady=(8, 0))

        # Resume PDF controls row
        self.resume_label = ctk.CTkLabel(controls, text="Resume (PDF)")
        self.resume_label.grid(row=3, column=0, sticky="w", pady=(8, 0))
        self.resume_entry = ctk.CTkEntry(controls, textvariable=self.resume_path, placeholder_text="Select your resume PDF…")
        self.resume_entry.grid(row=3, column=1, columnspan=6, sticky="ew", padx=6, pady=(8, 0))
        self.resume_browse_btn = ctk.CTkButton(controls, text="Browse", command=self._browse_resume)
        self.resume_browse_btn.grid(row=3, column=7, sticky="ew", pady=(8, 0))
        self.resume_clear_btn = ctk.CTkButton(controls, text="Clear", command=self._clear_resume_path)
        self.resume_clear_btn.grid(row=3, column=8, sticky="ew", padx=(6, 0), pady=(8, 0))

        # Email subject template row
        ctk.CTkLabel(controls, text="Subject Template (use {firm}, {school}, etc.)").grid(row=4, column=0, sticky="w", pady=(8, 0))
        self.subj_entry = ctk.CTkEntry(controls, textvariable=self.subject_template,
                                    placeholder_text="e.g. Duke Student interested in IB at {firm}")
        self.subj_entry.grid(row=4, column=1, columnspan=10, sticky="ew", padx=6, pady=(8, 0))
        self.subj_entry.bind("<FocusOut>", lambda e: self._save_settings())

        # --------- MIDDLE SECTION: Templates and Editor ----------
        mid = ctk.CTkFrame(main, fg_color="transparent")
        mid.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        mid.grid_columnconfigure(0, weight=0)  # Templates panel (fixed width)
        mid.grid_columnconfigure(1, weight=1)  # Editor panel (expandable)
        mid.grid_rowconfigure(1, weight=1)

        # Templates section label
        ctk.CTkLabel(mid, text="Templates").grid(row=0, column=0, sticky="w")

        # Left panel: Template list and management buttons
        left = ctk.CTkFrame(mid)
        left.grid(row=1, column=0, sticky="nsw", padx=(0, 12))
        left.grid_rowconfigure(0, weight=1)
        left.grid_columnconfigure(0, weight=0)

        # Template listbox with monospace font for consistent formatting
        self.templates_font = tkfont.Font(family=self.listbox_font_family, size=self.listbox_font_size)
        self.template_listbox = tk.Listbox(
            left, width=self.listbox_char_width, height=self.listbox_height,
            font=self.templates_font, activestyle="dotbox", exportselection=False
        )
        self.template_listbox.grid(row=0, column=0, sticky="nsw")

        # Scrollbar for template list
        sb = tk.Scrollbar(left, orient="vertical", command=self.template_listbox.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.template_listbox.configure(yscrollcommand=sb.set)
        self.template_listbox.bind("<<ListboxSelect>>", self._on_template_select)

        # Template management buttons
        left_btns = ctk.CTkFrame(mid, fg_color="transparent")
        left_btns.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        left_btns.grid_columnconfigure(0, weight=1)

        # Template operation buttons (stacked vertically)
        ctk.CTkButton(left_btns, text="New Template", command=self._new_template).grid(row=0, column=0, sticky="ew", pady=4)
        ctk.CTkButton(left_btns, text="Rename (F2)", command=self._rename_template).grid(row=1, column=0, sticky="ew", pady=4)
        ctk.CTkButton(left_btns, text="Move Up ▲", command=self._move_up).grid(row=2, column=0, sticky="ew", pady=4)
        ctk.CTkButton(left_btns, text="Move Down ▼", command=self._move_down).grid(row=3, column=0, sticky="ew", pady=4)
        ctk.CTkButton(left_btns, text="Save (⌘/Ctrl+S)", command=self._save_template_from_editor).grid(row=4, column=0, sticky="ew", pady=4)
        ctk.CTkButton(left_btns, text="Revert", command=self._revert_current).grid(row=5, column=0, sticky="ew", pady=4)
        ctk.CTkButton(left_btns, text="Remove", fg_color="#7f1d1d", hover_color="#991b1b",
                      command=self._remove_template).grid(row=6, column=0, sticky="ew", pady=4)
        ctk.CTkButton(left_btns, text="Import .txt Templates", command=self._import_txts).grid(row=7, column=0, sticky="ew", pady=4)
        ctk.CTkButton(left_btns, text="Export .zip to Downloads", command=self._export_templates_zip).grid(row=8, column=0, sticky="ew", pady=4)

        # Right panel: Template editor
        right = ctk.CTkFrame(mid)
        right.grid(row=0, column=1, rowspan=3, sticky="nsew")
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(2, weight=1)

        # Editor header with title and unsaved indicator
        header = ctk.CTkFrame(right, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 0))

        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=0)
        header.grid_columnconfigure(2, weight=0)

        self.editor_title = ctk.CTkLabel(header, text="Template Editor")
        self.editor_title.grid(row=0, column=0, sticky="w")

        self.unsaved_label = ctk.CTkLabel(header, text="", text_color="#f87171")
        self.unsaved_label.grid(row=0, column=1, sticky="e", padx=(8, 0))

        self.manual_only_var = tk.BooleanVar(value=False)

        self.manual_only_cb = ctk.CTkCheckBox(
            header,
            text="Manual-only (exclude from rotation)",
            variable=self.manual_only_var,
            command=self._on_manual_only_toggled
        )
        self.manual_only_cb.grid(row=0, column=2, sticky="e", padx=(12, 0))


        # Placeholder help text
        ctk.CTkLabel(
            right,
            text="Placeholders: {first name}, {last name}, {full name}, {firm}, {firm name}, {school}, or any exact CSV header.",
            text_color=("#444", "#aaa")
        ).grid(row=1, column=0, sticky="w", padx=4)

        # Main text editor for template content
        self.editor = ctk.CTkTextbox(right, wrap="word")
        self.editor.grid(row=2, column=0, sticky="nsew", padx=4, pady=6)
        self.editor.bind("<<Modified>>", self._on_editor_modified)

        # --------- BOTTOM SECTION: Action Buttons ----------
        bottom = ctk.CTkFrame(main, fg_color="transparent")
        bottom.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        bottom.grid_columnconfigure((0, 1, 2, 3), weight=1)

        # Main action buttons (equal width)
        self.btn_preview = ctk.CTkButton(bottom, text="Preview Recipients", command=self._handle_preview_clicked)
        self.btn_preview.grid(row=0, column=0, sticky="ew", padx=4)

        self.btn_generate = ctk.CTkButton(bottom, text="Generate Emails (Drafts)", command=self._handle_generate_clicked)
        self.btn_generate.grid(row=0, column=1, sticky="ew", padx=4)

        self.btn_clear_all = ctk.CTkButton(bottom, text="Clear All", command=self._clear_all)
        self.btn_clear_all.grid(row=0, column=2, sticky="ew", padx=4)

        self.btn_save_settings = ctk.CTkButton(bottom, text="Save Settings", command=self._save_settings)
        self.btn_save_settings.grid(row=0, column=3, sticky="ew", padx=4)

        # ========== POST-UI INITIALIZATION ==========
        
        # Apply the active profile's data to UI elements
        active = self.active_profile.get()
        prof = self.profiles.get(active, self._default_profile())
        self._apply_profile_state(prof)

        # Update UI state based on loaded data
        self._update_profile_menu_values()
        self.profile_menu.set(active)
        self._apply_source_enabled_state()
        
        # Select first template if available
        if self.templates and self.current_index is None:
            self._select_template(0)
        else:
            self._refresh_template_list()

    # =========================
    # Data Source Management
    # =========================
    def _on_change_data_source(self, label: str):
        """
        Handle data source toggle between Google Sheets and Local CSV.
        Updates internal state and UI element availability.
        """
        mode = "sheet" if label == "Google Sheet" else "csv"
        self.data_source.set(mode)
        self._apply_source_enabled_state()
        self._save_settings()
        self.toast.show(f"Using {label} as data source.", kind="info")

    def _set_enabled(self, widget, enabled: bool):
        """
        Safely set widget enabled/disabled state.
        Handles exceptions gracefully for widgets that might not support state changes.
        """
        try:
            widget.configure(state="normal" if enabled else "disabled")
        except Exception:
            pass  # Some widgets don't support state configuration

    def _apply_source_enabled_state(self):
        """
        Enable/disable UI controls based on selected data source.
        Google Sheet controls are enabled when using sheets, CSV controls when using CSV.
        Resume controls are always available regardless of data source.
        """
        using_sheet = self.data_source.get() == "sheet"
        
        # Skip if UI hasn't been built yet
        if not hasattr(self, "gsheet_entry"):
            return
        
        # Google Sheet controls (enabled only when using sheets)
        self._set_enabled(self.gsheet_entry, using_sheet)
        self._set_enabled(self.gsheet_load_btn, using_sheet)
        self._set_enabled(self.gsheet_clear_btn, using_sheet)
        self._set_enabled(self.gsheet_link_btn, using_sheet)
        
        # CSV controls (enabled only when using CSV)
        self._set_enabled(self.csv_entry, not using_sheet)
        self._set_enabled(self.csv_browse_btn, not using_sheet)
        self._set_enabled(self.csv_clear_btn, not using_sheet)
        
        # Resume controls are always enabled (independent of data source)

    def _ensure_latest_data(self) -> bool:
        if self.data_source.get() == "sheet":
            if not self.gsheet_url.get().strip():
                self.toast.show("Please paste a Google Sheets URL.", kind="error")
                return False
            return True
        else:
            csv_path = self.csv_path.get().strip()
            if not csv_path or not os.path.isfile(csv_path):
                self.toast.show("Please select a valid CSV file.", kind="error")
                return False
            return True


    # =========================
    # Google Sheets Integration
    # =========================
    GS_HOST = "docs.google.com"  # Google Sheets hostname for URL validation


    
    def _open_license_modal(self):
        """
        Display the license management modal window.
        
        Provides interface for:
        - Viewing current license status
        - Entering/editing license key
        - Validating license with the server
        """
        # Create modal window
        win = ctk.CTkToplevel(self)
        win.title("License")
        win.geometry("440x220")
        win.minsize(420, 200)
        win.transient(self)  # Stay on top of main window
        win.grab_set()  # Modal behavior

        # Main container
        wrapper = ctk.CTkFrame(win, fg_color="transparent")
        wrapper.pack(fill="both", expand=True, padx=16, pady=16)
        wrapper.grid_columnconfigure(0, weight=1)

        # Header
        ctk.CTkLabel(
            wrapper,
            text="License",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, sticky="w")

        # Status indicator (updates based on current license state)
        status_lbl = ctk.CTkLabel(wrapper, text="")
        status_lbl.grid(row=1, column=0, sticky="w", pady=(6, 10))

        def _refresh_status():
            """Update status display based on current license state."""
            if getattr(self, "_licensed", False):
                status_lbl.configure(text="✓ Licensed", text_color="#10b981")
            else:
                status_lbl.configure(text="● Unlicensed", text_color="#f87171")

        _refresh_status()

        # License key input field
        self.license_entry = ctk.CTkEntry(
            wrapper,
            textvariable=self.license_key,
            placeholder_text="Enter your license key…"
        )
        self.license_entry.grid(row=2, column=0, sticky="ew")
        self.license_entry.focus_set()  # Auto-focus for immediate typing

        # Button row (right-aligned)
        btnrow = ctk.CTkFrame(wrapper, fg_color="transparent")
        btnrow.grid(row=3, column=0, sticky="e", pady=(14, 0))
        btnrow.grid_columnconfigure(0, weight=1)

        def _do_validate():
            """Handle license validation button click."""
            self._validate_license()   # Update app-wide license state
            self._save_settings()      # Persist license key
            _refresh_status()          # Update modal display

        # Action buttons
        ctk.CTkButton(btnrow, text="Validate", command=_do_validate).grid(row=0, column=1, padx=(0, 8))
        ctk.CTkButton(btnrow, text="Close", command=win.destroy).grid(row=0, column=2)

        # Enable Enter key to validate
        self.license_entry.bind("<Return>", lambda e: _do_validate())

    

    def _load_gsheet(self, quiet: bool = True):
        """
        Google Sheets loading is handled by engine.data_sources.
        This method is kept only to satisfy UI flow.
        """
        if not quiet:
            self.toast.show("Google Sheet will be loaded when preview or generate runs.", kind="info")


    def _clear_gsheet_url(self):
        """Clear Google Sheets URL and cached data."""
        self.gsheet_url.set("")
        self._gsheet_rows_cache = None
        self._last_headers_lower = []
        self._save_settings()
        self.toast.show("Cleared Google Sheets URL.", kind="info")

    def _open_gsheet_in_browser(self):
        """Open the current Google Sheets URL in the default web browser."""
        url = self.gsheet_url.get().strip()
        if not url:
            self.toast.show("No Google Sheets URL to open.", kind="warning")
            return
        try:
            webbrowser.open(url)
        except Exception as e:
            self.toast.show(f"Couldn't open URL: {e}", kind="error")

    # =========================
    # Data Reading and Processing
    # =========================
    def _read_all_rows(self):
        """
        Read all rows from the selected data source using engine loaders.
        """
        if self.data_source.get() == "sheet":
            rows, headers = load_google_sheet(self.gsheet_url.get())
        else:
            rows, headers = load_csv(self.csv_path.get())

        self._last_headers_lower = headers
        return rows


    def _read_eligible_rows(self):
        """
        Read only rows eligible for email generation.
        
        Filters rows to include only those with:
        - Generate column = True/1/Yes
        - Valid email address
        
        Returns:
            list: Filtered rows ready for email generation
        """
        rows = self._read_all_rows()
        out = []
        
        for r in rows:
            resolver = self._make_resolver(r)
            email = resolver.get_email()
            
            # Include only rows marked for generation with valid emails
            if self._is_generate_true(r) and self._is_email_valid(email):
                out.append(r)
        
        return out


    

    # =========================
    # Preview Window Management
    # =========================
    def _handle_preview_clicked(self):
        """
        Handle the Preview Recipients button click.
        Ensures data is current before opening preview window.
        """
        if not self._ensure_latest_data():
            return
        self._open_preview_window()

    def _close_preview_window(self):
        """Safely close the preview window if it exists."""
        try:
            if self._preview_win and tk.Toplevel.winfo_exists(self._preview_win):
                self._preview_win.destroy()
        except Exception:
            pass  # Handle case where window was already destroyed
        self._preview_win = None

    def _set_inspector_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        self._insp_assignment_menu.configure(state=state)
        self._insp_clear_btn.configure(state=state)

        if not enabled:
            self._insp_name.configure(text="—")
            self._insp_email.configure(text="")
            self._insp_firm.configure(text="—")
            self._insp_assignment_var.set("Auto (rotation)")


    def _open_preview_window(self):
        """
        Open or bring to front the recipient preview window.
        
        Shows a table of recipients with their extracted information
        and which template would be used for each.
        """
        # If window already exists, just bring it to front and refresh
        if self._preview_win and tk.Toplevel.winfo_exists(self._preview_win):
            self._preview_win.lift()
            self._populate_preview_table()
            return

        # Create new preview window
        self._preview_win = ctk.CTkToplevel(self)
        self._preview_win.title("Preview Recipients")
        self._preview_win.geometry("900x520")
        self._preview_win.minsize(1400, 420)
        self._preview_win.transient(self)

        # Header controls
        header = ctk.CTkFrame(self._preview_win, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(12, 6))

        # Toggle for showing all rows vs only recipients
        self._preview_only_recipients.set(True)
        switch = ctk.CTkSwitch(
            header,
            text="Show only recipients (Generate = True & valid email)",
            variable=self._preview_only_recipients,
            command=self._populate_preview_table
        )
        switch.pack(side="left")

        # Row count display
        self._preview_count_lbl = ctk.CTkLabel(header, text="")
        self._preview_count_lbl.pack(side="right")

        # Table container
        content = ctk.CTkFrame(self._preview_win)
        content.pack(fill="both", expand=True, padx=12, pady=12)

        content.grid_columnconfigure(0, weight=3)
        content.grid_columnconfigure(1, weight=2)

        table_frame = ttk.Frame(content)
        table_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 12))


        # Create treeview table with columns
        self.preview_tree = ttk.Treeview(
            table_frame,
            columns=("name", "email", "firm", "template"),
            show="headings",  # Hide the tree column
            height=16,
        )
        
        # Configure column headers
        self.preview_tree.heading("name", text="Name")
        self.preview_tree.heading("email", text="Email")
        self.preview_tree.heading("firm", text="Firm")
        self.preview_tree.heading("template", text="Template")

        # Configure column widths and alignment
        self.preview_tree.column("name", width=220, anchor="w")
        self.preview_tree.column("email", width=260, anchor="w")
        self.preview_tree.column("firm", width=220, anchor="w")
        self.preview_tree.column("template", width=160, anchor="w")

        # Add table and scrollbar
        self.preview_tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(table_frame, orient="vertical", command=self.preview_tree.yview)
        sb.pack(side="right", fill="y")
        self.preview_tree.configure(yscrollcommand=sb.set)
        self.preview_tree.bind("<<TreeviewSelect>>", self._on_preview_selection_changed)
        self._build_preview_inspector(content)
        # Populate initial data
        self._populate_preview_table()

    def _build_preview_inspector(self, parent):
        panel = ctk.CTkFrame(parent)
        panel.grid(row=0, column=1, sticky="nsew")

        self._preview_inspector_panel = panel

        ctk.CTkLabel(panel, text="Recipient", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", pady=(8, 4))

        self._insp_name = ctk.CTkLabel(panel, text="—")
        self._insp_name.pack(anchor="w")

        self._insp_email = ctk.CTkLabel(panel, text="", text_color="gray")
        self._insp_email.pack(anchor="w", pady=(0, 8))

        ctk.CTkLabel(panel, text="Firm").pack(anchor="w")
        self._insp_firm = ctk.CTkLabel(panel, text="—")
        self._insp_firm.pack(anchor="w", pady=(0, 12))

        ctk.CTkLabel(panel, text="Template Assignment").pack(anchor="w")

        self._insp_assignment_var = tk.StringVar(value="Auto (rotation)")

        self._insp_assignment_menu = ctk.CTkOptionMenu(
            panel,
            variable=self._insp_assignment_var,
            values=["Auto (rotation)"],
            command=lambda _: self._on_preview_assignment_changed(),
            width=260,
        )
        self._insp_assignment_menu.pack(anchor="w", pady=(4, 12))

        self._insp_clear_btn = ctk.CTkButton(
            panel,
            text="Clear override",
            command=lambda: (
                self._insp_assignment_var.set("Auto (rotation)"),
                self._on_preview_assignment_changed()
            )
        )

        self._insp_clear_btn.pack(anchor="w")

        self._set_inspector_enabled(False)

    def _on_preview_selection_changed(self, _evt=None):
        sel = self.preview_tree.selection()
        if not sel:
            self._set_inspector_enabled(False)
            return

        item = sel[0]

        email = (self._preview_item_to_email.get(item) or "").lower().strip()
        if not self._is_email_valid(email):
            self._set_inspector_enabled(False)
            return

        name, _, firm, _ = self.preview_tree.item(item, "values")


        if not self._is_email_valid(email):
            self._set_inspector_enabled(False)
            return

        self._preview_selected_email = email

        self._insp_name.configure(text=name)
        self._insp_email.configure(text=email)
        self._insp_firm.configure(text=firm or "—")

        menu_values = ["Auto (rotation)"]
        for t in self.templates:
            label = t["name"]
            if t.get("manual_only"):
                label += " (manual-only)"
            menu_values.append(label)

        self._insp_assignment_menu.configure(values=menu_values)

        if email in self.recipient_template_overrides:
            t = self._template_by_id(self.recipient_template_overrides[email])
            if t:
                label = t["name"]
                if t.get("manual_only"):
                    label += " (manual-only)"
                self._insp_assignment_var.set(label)
            else:
                self._insp_assignment_var.set("Auto (rotation)")
        else:
            self._insp_assignment_var.set("Auto (rotation)")

        self._set_inspector_enabled(True)

    def _on_preview_assignment_changed(self):
        email = getattr(self, "_preview_selected_email", None)
        if not email:
            return

        choice = self._insp_assignment_var.get()

        if choice == "Auto (rotation)":
            self.recipient_template_overrides.pop(email, None)
        else:
            clean = choice.replace(" (manual-only)", "")
            for t in self.templates:
                if t["name"] == clean:
                    self.recipient_template_overrides[email] = t["id"]
                    break

        self._save_settings()

        self._populate_preview_table(reselect_email=email)




    def _rotatable_templates(self):
        return [t for t in self.templates if not t.get("manual_only", False)]
    

    def _template_by_id(self, tid):
        for t in self.templates:
            if t.get("id") == tid:
                return t
        return None
    
    def _prune_recipient_overrides(self):
        valid_ids = {t.get("id") for t in self.templates}
        dead = [email for email, tid in (self.recipient_template_overrides or {}).items() if tid not in valid_ids]
        for email in dead:
            self.recipient_template_overrides.pop(email, None)

        
    def _populate_preview_table(self, reselect_email: str | None = None):
        """
        Populate the preview table with recipient data and
        optionally reselect a previously selected recipient.
        """

        # --- remember who was selected BEFORE rebuild ---
        prev_email = reselect_email or getattr(self, "_preview_selected_email", None)

        # --- clear table ---
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)

        self._preview_item_to_email = {}
        self._preview_selected_email = None
        self._set_inspector_enabled(False)

        only_recipients = self._preview_only_recipients.get()

        # ==================================================
        # BUILD PREVIEW ROWS VIA ENGINE
        # ==================================================
        if only_recipients:
            rows = self._read_eligible_rows()
        else:
            rows = [r for r in self._read_all_rows() if self._has_email_populated(r)]

        preview_rows = build_preview_rows(
            rows=rows,
            headers_lower=self._last_headers_lower,
            templates=self.templates,
            recipient_template_overrides=self.recipient_template_overrides,
            parse_name_fn=self._parse_name,
            only_recipients=only_recipients,
            is_generate_true_fn=self._is_generate_true,
            is_email_valid_fn=self._is_email_valid,
        )

        # ==================================================
        # RENDER PREVIEW ROWS
        # ==================================================
        for pr in preview_rows:
            tpl = pr["template_name"]
            if pr.get("is_manual"):
                tpl += " (manual)"

            item_id = self.preview_tree.insert(
                "",
                "end",
                values=(pr["name"], pr["email"], pr["firm"], tpl)
            )

            # stable identity mapping
            self._preview_item_to_email[item_id] = pr["email"]

        self._preview_count_lbl.configure(text=f"Total: {len(preview_rows)}")

        # ==================================================
        # RESELECT PREVIOUS ROW (THIS IS STEP 3)
        # ==================================================
        if prev_email:
            for item_id, email in self._preview_item_to_email.items():
                if email == prev_email:
                    self.preview_tree.selection_set(item_id)
                    self.preview_tree.see(item_id)

                    self._preview_selected_email = prev_email
                    self._on_preview_selection_changed()
                    break

    
    # def _template_name_for_next_in_firm(self, firm_counts, firm):
    #     """
    #     Determine which template would be used for the next person at a firm.
        
    #     Templates rotate within each firm to provide variety.
    #     Uses a firm_counts dictionary to track template rotation.
        
    #     Args:
    #         firm_counts: Dictionary tracking how many people per firm
    #         firm: The firm name (empty string for no firm)
            
    #     Returns:
    #         str: Template name that would be used
    #     """
    #     if not self.templates:
    #         return ""
        
    #     firm_key = firm or ""
    #     firm_counts[firm_key] = firm_counts.get(firm_key, 0) + 1
    #     idx = (firm_counts[firm_key] - 1) % len(self.templates)
    #     return self.templates[idx]["name"]

    # =========================
    # Settings Persistence (JSON with Profiles)
    # =========================

    def _destroy_preview_cell_editor(self):
        try:
            if getattr(self, "_preview_cell_editor", None) is not None:
                self._preview_cell_editor.destroy()
        except Exception:
            pass
        self._preview_cell_editor = None

    def _on_preview_tree_single_click(self, event):
        row_id = self.preview_tree.identify_row(event.y)
        col_id = self.preview_tree.identify_column(event.x)

        # Only act on Template column
        if col_id == "#4":
            self._on_preview_tree_double_click(event)
    

    def _on_preview_tree_double_click(self, event):
        row_id = self.preview_tree.identify_row(event.y)
        col_id = self.preview_tree.identify_column(event.x)

        # Template column is the 4th column: ("name","email","firm","template")
        if not row_id or col_id != "#4":
            return

        email = (self._preview_item_to_email.get(row_id) or "").strip().lower()
        if not email or not self._is_email_valid(email):
            return

        bbox = self.preview_tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        self._destroy_preview_cell_editor()

        values = ["Auto (rotation)"] + [t.get("name", "") for t in self.templates]

        # Preselect current value
        current_name = "Auto (rotation)"
        if email in self.recipient_template_overrides:
            t = self._template_by_id(self.recipient_template_overrides[email])
            if t:
                current_name = t.get("name", "Auto (rotation)")
        else:
            row_vals = self.preview_tree.item(row_id, "values") or ()
            if len(row_vals) >= 4:
                cell = str(row_vals[3]).replace(" (manual)", "").strip()
                if cell and cell != "–":
                    current_name = cell

        var = tk.StringVar(value=current_name if current_name in values else "Auto (rotation)")
        style = ttk.Style()
        style.configure(
            "Preview.TCombobox",
            padding=6,
        )

        cb = ttk.Combobox(
            self.preview_tree.master,
            textvariable=var,
            values=values,
            state="readonly",
            style="Preview.TCombobox"
        )

        
        tree_x = self.preview_tree.winfo_x()
        tree_y = self.preview_tree.winfo_y()

        cb.place(
            x=tree_x + x,
            y=tree_y + y,
            width=w,
            height=h
        )


        cb.focus_set()

        def _commit(_evt=None):
            choice = var.get()

            if choice == "Auto (rotation)":
                self.recipient_template_overrides.pop(email, None)
            else:
                chosen_id = None
                for t in self.templates:
                    if t.get("name") == choice:
                        chosen_id = t.get("id")
                        break
                if chosen_id:
                    self.recipient_template_overrides[email] = chosen_id

            self._save_settings()
            self._populate_preview_table()

        cb.bind("<<ComboboxSelected>>", _commit)
        cb.bind("<Return>", _commit)
        cb.bind("<FocusOut>", lambda e: self._destroy_preview_cell_editor())

        self._preview_cell_editor = cb

    def _default_profile(self):
        """
        Create a default profile with empty/default values.
        
        Returns:
            dict: Default profile configuration
        """
        return {
            "csv_path": "",
            "resume_path": "",
            "subject_template": DEFAULT_SUBJECT,
            "recipient_template_overrides": {},  # email_lower -> template_id
            "gsheet_url": "",
            "data_source": "sheet",
            "templates": [],
        }

    def _load_settings(self):
        """
        Load application settings from disk.
        
        Handles:
        - Missing settings file (creates defaults)
        - Schema migration from older versions
        - Profile validation and normalization
        - Global license key management
        """
        # Create default state if no settings file exists
        if not os.path.exists(SETTINGS_FILE):
            self.profiles = {"Default": self._default_profile()}
            self.profile_order = ["Default"]
            self.active_profile.set("Default")
            self.license_key.set("")
            return

        # Load and parse settings file
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            # Fallback to defaults if file is corrupted
            self.profiles = {"Default": self._default_profile()}
            self.profile_order = ["Default"]
            self.active_profile.set("Default")
            self.license_key.set("")
            return

        # Handle modern multi-profile format
        if "profiles" in data and "active_profile" in data:
            self.profiles = data.get("profiles") or {"Default": self._default_profile()}
            
            # Normalize and validate all profiles
            for k, v in list(self.profiles.items()):
                if not isinstance(v, dict):
                    self.profiles[k] = self._default_profile()
                else:
                    # Ensure all required keys exist with defaults
                    v.setdefault("csv_path", "")
                    v.setdefault("resume_path", "")
                    v.setdefault("subject_template", DEFAULT_SUBJECT)
                    v.setdefault("gsheet_url", "")
                    v.setdefault("data_source", "sheet")
                    v.setdefault("templates", [])
                    v.setdefault("recipient_template_overrides", {})
                    
                    # Remove deprecated per-profile license keys
                    if "license_key" in v:
                        v.pop("license_key", None)

            # Validate and rebuild profile order
            self.profile_order = data.get("profile_order") or list(self.profiles.keys())
            self.profile_order = [n for n in self.profile_order if n in self.profiles]
            
            # Add any missing profiles to the order
            for name in self.profiles.keys():
                if name not in self.profile_order:
                    self.profile_order.append(name)
            
            # Validate active profile
            act = data.get("active_profile")
            if not act or act not in self.profiles:
                act = self.profile_order[0]
            self.active_profile.set(act)

            # Load global license key (migrate from old per-profile storage if needed)
            lk = data.get("license_key", "")
            if not lk:
                # Migration: find license key from any old profile
                for v in self.profiles.values():
                    old = v.pop("license_key", "")
                    if old:
                        lk = old
                        break
            self.license_key.set(lk or "")

        else:
            # Migration from very old single-profile format
            prof = {
                "csv_path": data.get("csv_path", ""),
                "resume_path": data.get("resume_path", ""),
                "subject_template": data.get("subject_template", DEFAULT_SUBJECT),
                "gsheet_url": data.get("gsheet_url", ""),
                "data_source": data.get("data_source", "sheet"),
                "templates": self._migrate_templates(data.get("templates", [])),
            }
            self.profiles = {"Default": prof}
            self.profile_order = ["Default"]
            self.active_profile.set("Default")
            self.license_key.set(data.get("license_key", ""))
            
            # Save migrated format
            self._save_settings()

    def _atomic_save_json(self, path: str, data: dict):
        """
        Atomically save JSON data to prevent corruption.
        
        Uses a temporary file in the same directory, writes completely,
        then atomically moves it to the final location. This prevents
        data loss if the program crashes during saving.
        
        Args:
            path: Destination file path
            data: Dictionary to save as JSON
        """
        p = Path(path).expanduser()
        parent = p.parent
        tmp_path = None

        try:
            # Ensure destination directory exists
            parent.mkdir(parents=True, exist_ok=True)

            # Write to temporary file in same directory (required for atomic move)
            with tempfile.NamedTemporaryFile(
                mode="w",
                encoding="utf-8",
                dir=str(parent),
                prefix=p.name + ".",
                suffix=".tmp",
                delete=False,
            ) as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
                f.flush()
                os.fsync(f.fileno())  # Force write to disk
                tmp_path = Path(f.name)

            # Atomic replacement (works on all platforms)
            os.replace(str(tmp_path), str(p))

            # Best-effort directory sync for durability (POSIX only)
            try:
                dir_fd = os.open(str(parent), getattr(os, "O_DIRECTORY", 0))
                try:
                    os.fsync(dir_fd)
                finally:
                    os.close(dir_fd)
            except Exception:
                pass  # Not critical; ignore on unsupported platforms

        except Exception:
            # Clean up temporary file if anything failed
            try:
                if tmp_path and tmp_path.exists():
                    tmp_path.unlink()
            except Exception:
                pass
            raise

    def _save_settings(self):
        """
        Save current application state to disk.
        
        Collects current profile state and saves all profiles,
        profile order, active profile, and global settings.
        """
        name = self.active_profile.get() or "Default"
        
        # Ensure current profile exists
        if name not in self.profiles:
            self.profiles[name] = self._default_profile()
            if name not in self.profile_order:
                self.profile_order.append(name)
        
        # Update current profile with UI state
        self.profiles[name] = self._collect_current_profile_state()
        
        # Build complete settings data
        data = {
            "schema_version": SCHEMA_VERSION,
            "active_profile": name,
            "profile_order": self.profile_order,
            "profiles": self.profiles,
            "license_key": self.license_key.get(),  # Global license key
        }
        
        # Save atomically to prevent corruption
        try:
            os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
            self._atomic_save_json(SETTINGS_FILE, data)
        except Exception:
            print("Failed to save settings JSON.", file=sys.stderr)
            self.toast.show("Failed to save settings to disk.", kind="error")

    # =========================
    # Profile Management System
    # =========================
    def _update_profile_menu_values(self):
        """Update the profile dropdown menu with current profile list."""
        vals = self.profile_order if self.profile_order else ["Default"]
        try:
            self.profile_menu.configure(values=vals)
        except Exception:
            pass  # Handle case where menu doesn't exist yet

    def _validate_profile_name(self, name: str, allow_same_case_change=False):
        """
        Validate a profile name for creation or renaming.
        
        Checks for:
        - Non-empty name
        - Reasonable length
        - Case-insensitive uniqueness
        
        Args:
            name: Proposed profile name
            allow_same_case_change: Allow changing case of current profile name
            
        Returns:
            tuple: (is_valid, error_message)
        """
        if name is None:
            return False, "Name is required."
        
        n = name.strip()
        if not n:
            return False, "Name is required."
        
        if len(n) > 64:
            return False, "Name is too long (max 64 characters)."
        
        # Check for case-insensitive duplicates
        cf = n.casefold()
        current = self.active_profile.get()
        
        for existing in self.profiles.keys():
            # Skip current profile if allowing case changes
            if allow_same_case_change and existing == current:
                continue
            if existing.casefold() == cf:
                return False, "A profile with that name already exists."
        
        return True, None

    def _commit_all_unsaved_buffers(self):
        """
        Save all unsaved template changes to their respective templates.
        Clears the unsaved buffers after committing.
        """
        if not self.unsaved_buffers:
            return
        
        # Apply each unsaved change to its template
        for idx, text in list(self.unsaved_buffers.items()):
            if 0 <= idx < len(self.templates):
                self.templates[idx]["text"] = text or ""
        
        # Clear all unsaved indicators
        self.unsaved_buffers.clear()
        self._update_editor_header()

    def _collect_current_profile_state(self) -> dict:
        """
        Collect current UI state into a profile dictionary.
        
        Returns:
            dict: Current profile state with all settings
        """
        return {
            "csv_path": self.csv_path.get(),
            "resume_path": self.resume_path.get(),
            "subject_template": self.subject_template.get(),
            "gsheet_url": self.gsheet_url.get(),
            "data_source": self.data_source.get(),
            "templates": self.templates,
            "recipient_template_overrides": self.recipient_template_overrides,
        }


    def _apply_profile_state(self, profile: dict):
        """
        Apply a profile's settings to the UI.
        
        Updates all UI elements and clears unsaved changes.
        
        Args:
            profile: Profile dictionary with settings to apply
        """
        # Update UI variables
        self.csv_path.set(profile.get("csv_path", ""))
        self.resume_path.set(profile.get("resume_path", ""))
        self.subject_template.set(profile.get("subject_template", DEFAULT_SUBJECT))
        self.gsheet_url.set(profile.get("gsheet_url", ""))
        self.data_source.set(profile.get("data_source", "sheet"))
        
        # Update template and editor state
        self.templates = self._migrate_templates(profile.get("templates", []))
        self.recipient_template_overrides = dict(profile.get("recipient_template_overrides", {}) or {})
        self.unsaved_buffers = {}
        self.current_index = None

        # Update UI controls
        try:
            self.source_seg.set("Google Sheet" if self.data_source.get() == "sheet" else "Local CSV")
        except Exception:
            pass  # Handle case where UI doesn't exist yet
        
        self._apply_source_enabled_state()
        self._refresh_template_list()
        self.editor.delete("1.0", "end")
        self._update_editor_header()
        
        # Clear cached data for new profile
        self._gsheet_rows_cache = None
        self._last_headers_lower = []

    def _unsaved_changes_present(self) -> bool:
        """
        Check if there are any unsaved changes in the current profile.
        
        Compares current UI state with saved profile data.
        
        Returns:
            bool: True if there are unsaved changes
        """
        # Check template unsaved buffers
        if self.unsaved_buffers:
            return True
        
        # Compare current UI state with saved profile
        name = self.active_profile.get()
        prof = self.profiles.get(name, self._default_profile())
        
        # Check each field for changes
        if self.csv_path.get() != prof.get("csv_path", ""):
            return True
        if self.resume_path.get() != prof.get("resume_path", ""):
            return True
        if self.subject_template.get() != prof.get("subject_template", DEFAULT_SUBJECT):
            return True
        if self.gsheet_url.get() != prof.get("gsheet_url", ""):
            return True
        if self.data_source.get() != prof.get("data_source", "sheet"):
            return True
        
        return False

    def _prompt_unsaved_and_resolve(self, intent_label: str) -> str:
        """
        Prompt user about unsaved changes and get their decision.
        
        Args:
            intent_label: Description of the action being attempted
            
        Returns:
            str: User decision - "save", "discard", or "cancel"
        """
        if not self._unsaved_changes_present():
            return "save"  # No changes, safe to proceed
        
        # Show save/discard/cancel dialog
        resp = messagebox.askyesnocancel(
            title="Unsaved changes",
            message=f"You have unsaved changes in \"{self.active_profile.get()}\". Save before {intent_label}?",
            icon="warning",
        )
        
        if resp is None:
            return "cancel"
        return "save" if resp else "discard"

    def _on_profile_selected_ui(self, new_name: str):
        """
        Handle profile selection from the dropdown menu.
        
        Manages unsaved changes and switches to the new profile.
        
        Args:
            new_name: Name of the newly selected profile
        """
        # Ignore if selecting the same profile
        if new_name == self.active_profile.get():
            return
        
        # Handle unsaved changes
        decision = self._prompt_unsaved_and_resolve("switching profiles")
        if decision == "cancel":
            # Revert dropdown to current profile
            self.profile_menu.set(self.active_profile.get())
            return
        
        if decision == "save":
            # Save current profile state
            self._commit_all_unsaved_buffers()
            self.profiles[self.active_profile.get()] = self._collect_current_profile_state()
        else:
            # Discard changes
            self.unsaved_buffers.clear()

        # Switch to new profile
        self.active_profile.set(new_name)
        self._apply_profile_state(self.profiles[new_name])
        self._update_profile_menu_values()
        self.profile_menu.set(new_name)
        self._close_preview_window()  # Close preview since data may change
        self._save_settings()
        self.toast.show(f"Switched to profile \"{new_name}\".", kind="info")

    def _profile_new(self):
        """Create a new profile with a user-specified name."""
        name = simpledialog.askstring("New Profile", "Profile name:")
        if name is None:
            return
        
        # Validate the proposed name
        ok, msg = self._validate_profile_name(name)
        if not ok:
            self.toast.show(msg, kind="warning")
            return

        # Create the new profile
        self.profiles[name] = self._default_profile()
        self.profile_order.append(name)

        # Handle unsaved changes in current profile
        decision = self._prompt_unsaved_and_resolve("switching profiles")
        if decision == "cancel":
            # Save the new profile but don't switch to it
            self._save_settings()
            self._update_profile_menu_values()
            return
        
        if decision == "save":
            self._commit_all_unsaved_buffers()
            self.profiles[self.active_profile.get()] = self._collect_current_profile_state()
        else:
            self.unsaved_buffers.clear()

        # Switch to the new profile
        self.active_profile.set(name)
        self._apply_profile_state(self.profiles[name])
        self._update_profile_menu_values()
        self.profile_menu.set(name)
        self._close_preview_window()
        self._save_settings()
        self.toast.show(f"Created profile \"{name}\".", kind="success")

    def _profile_duplicate(self):
        """Duplicate the current profile with a new name."""
        src = self.active_profile.get()
        if not src or src not in self.profiles:
            return
        
        # Generate a unique name suggestion
        base = f"{src} Copy"
        new_name = base
        counter = 2
        existing_cf = {n.casefold() for n in self.profiles.keys()}
        
        while new_name.casefold() in existing_cf:
            new_name = f"{base} ({counter})"
            counter += 1
        
        # Get user input for the new name
        new_name = simpledialog.askstring("Duplicate Profile", "New profile name:", initialvalue=new_name)
        if not new_name:
            return
        
        # Validate the name
        ok, msg = self._validate_profile_name(new_name)
        if not ok:
            self.toast.show(msg, kind="warning")
            return

        # Save current state and create duplicate
        self._commit_all_unsaved_buffers()
        self.profiles[src] = self._collect_current_profile_state()

        # Deep copy the source profile
        src_prof = self.profiles[src]
        dup = {
            "csv_path": src_prof.get("csv_path", ""),
            "resume_path": src_prof.get("resume_path", ""),
            "subject_template": src_prof.get("subject_template", DEFAULT_SUBJECT),
            "gsheet_url": src_prof.get("gsheet_url", ""),
            "data_source": src_prof.get("data_source", "sheet"),
            "recipient_template_overrides": dict(src_prof.get("recipient_template_overrides", {}) or {}),
            "templates": [
                            {
                                "id": t.get("id") or str(uuid.uuid4()),
                                "name": t.get("name", ""),
                                "text": t.get("text", ""),
                                "manual_only": bool(t.get("manual_only", False)),
                            }
                            for t in src_prof.get("templates", [])
                        ],

        }
        
        self.profiles[new_name] = dup
        self.profile_order.append(new_name)

        # Switch to the duplicated profile
        self.active_profile.set(new_name)
        self._apply_profile_state(dup)
        self._update_profile_menu_values()
        self.profile_menu.set(new_name)
        self._close_preview_window()
        self._save_settings()
        self.toast.show(f"Duplicated to \"{new_name}\".", kind="success")

    def _profile_rename(self):
        """Rename the current profile."""
        old = self.active_profile.get()
        if not old:
            return
        
        # Get new name from user
        new_name = simpledialog.askstring("Rename Profile", f"New name for \"{old}\":", initialvalue=old)
        if not new_name or new_name == old:
            return
        
        # Validate the new name
        ok, msg = self._validate_profile_name(new_name, allow_same_case_change=True)
        if not ok:
            self.toast.show(msg, kind="warning")
            return

        # Save current state before renaming
        self._commit_all_unsaved_buffers()
        self.profiles[old] = self._collect_current_profile_state()

        # Rename the profile
        prof = self.profiles.pop(old)
        self.profiles[new_name] = prof
        
        # Update the profile order
        try:
            idx = self.profile_order.index(old)
            self.profile_order[idx] = new_name
        except ValueError:
            if new_name not in self.profile_order:
                self.profile_order.append(new_name)

        # Update active profile and UI
        self.active_profile.set(new_name)
        self._update_profile_menu_values()
        self.profile_menu.set(new_name)
        self._save_settings()
        self.toast.show(f"Renamed to \"{new_name}\".", kind="success")

    def _profile_delete(self):
        """Delete the current profile after confirmation."""
        name = self.active_profile.get()
        if not name or name not in self.profiles:
            return
        
        # Prevent deleting the last profile
        if len(self.profiles) <= 1:
            self.toast.show("At least one profile must exist.", kind="warning")
            return

        # Handle unsaved changes
        decision = self._prompt_unsaved_and_resolve("deleting this profile")
        if decision == "cancel":
            return
        
        if decision == "save":
            self._commit_all_unsaved_buffers()
            self.profiles[name] = self._collect_current_profile_state()
        else:
            self.unsaved_buffers.clear()

        # Confirm deletion
        if not messagebox.askokcancel("Delete Profile", f"Delete profile \"{name}\"? This cannot be undone."):
            return

        # Find the index for selecting the next profile
        try:
            idx = self.profile_order.index(name)
        except ValueError:
            idx = 0
        
        # Remove the profile
        self.profile_order = [n for n in self.profile_order if n != name]
        self.profiles.pop(name, None)

        # Select the next appropriate profile
        if self.profile_order:
            new_idx = max(0, min(idx - 1, len(self.profile_order) - 1))
            new_active = self.profile_order[new_idx]
        else:
            # Create a default profile if none remain
            new_active = "Default"
            self.profiles[new_active] = self._default_profile()
            self.profile_order = [new_active]

        # Switch to the new active profile
        self.active_profile.set(new_active)
        self._apply_profile_state(self.profiles[new_active])
        self._update_profile_menu_values()
        self.profile_menu.set(new_active)
        self._close_preview_window()
        self._save_settings()
        self.toast.show(f"Deleted \"{name}\".", kind="warning")
    def _open_help_modal(self):
        """
        Help window that reads README.md, converts Markdown -> clean, nicely spaced text.
        - Bigger window
        - Scrollable, read-only
        - Bullets look like bullets, headings get breathing room
        """
        import webbrowser
        from pathlib import Path
        import re
        import tkinter.font as tkfont
        from tkinter import messagebox

        # Third-party helpers (install if needed: pip install markdown beautifulsoup4)
        try:
            import markdown
            from bs4 import BeautifulSoup
            _md_available = True
        except Exception:
            _md_available = False

        # --- Window setup ---
        win = ctk.CTkToplevel(self)
        win.title("DraftMate — Help")
        win.geometry("980x720")
        win.minsize(820, 560)
        win.transient(self)
        win.grab_set()

        # --- Header ---
        header = ctk.CTkFrame(win, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(16, 8))

        ctk.CTkLabel(
            header, text="DraftMate — Help",
            font=ctk.CTkFont(size=20, weight="bold")
        ).pack(side="left")

        readme_path = Path(__file__).resolve().parent / "README.md"

        # Buttons
        def _open_template():
            webbrowser.open("https://docs.google.com/spreadsheets/d/1XsusJLv3BxSnaRLaHMGuaYN2DJbjp_TtmBNY8KzpxB0/edit?usp=sharing")

        def _copy_all():
            try:
                text = textbox.get("1.0", "end").strip()
                if not text:
                    return
                win.clipboard_clear()
                win.clipboard_append(text)
                messagebox.showinfo("Copied", "README content copied to clipboard.")
            except Exception:
                pass

        btns = ctk.CTkFrame(header, fg_color="transparent")
        btns.pack(side="right")
        ctk.CTkButton(btns, text="Open Template Sheet", command=_open_template).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btns, text="Copy", command=_copy_all).pack(side="left", padx=(0, 6))
        ctk.CTkButton(btns, text="Close", command=win.destroy).pack(side="left")

        # --- Body ---
        body = ctk.CTkFrame(win, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        textbox = ctk.CTkTextbox(body, wrap="word")
        textbox.pack(fill="both", expand=True)

        # Comfy font
        try:
            font = tkfont.Font(family="Helvetica", size=18)
            textbox.configure(font=font)
        except Exception:
            pass

        # --- Load README ---
        try:
            if readme_path.exists():
                raw_md = readme_path.read_text(encoding="utf-8")
            else:
                raw_md = (
                    "README.md not found next to the app.\n\n"
                    "Put README.md in the same folder as email_app.py, then reopen this Help window."
                )
        except Exception as e:
            raw_md = f"Couldn't read README.md:\n{e}"

        # --- Convert & pretty-format ---
        def _format_markdown_to_plain(md_text: str) -> str:
            """
            Turn Markdown into neat, human-friendly plain text:
            - Headings = single line with blank line after
            - Paragraphs = rejoined (no mid-sentence newlines)
            - Lists = '• item' per line, with a blank line after the list
            - Collapse extra blank lines
            """
            if not _md_available:
                # Fallback: very light cleanup if libs aren't installed
                t = md_text.replace("\r\n", "\n")
                # Strip common markers
                for a, b in [
                    ("**", ""), ("*", ""), ("__", ""), ("_", ""), ("`", ""),
                    ("### ", ""), ("## ", ""), ("# ", "")
                ]:
                    t = t.replace(a, b)
                # Normalize bullets
                t = re.sub(r"(?m)^\s*[-+*]\s+", "• ", t)
                # Collapse 3+ blank lines to one
                t = re.sub(r"\n{3,}", "\n\n", t)
                # Join wrapped lines inside paragraphs (two+ newlines keep paragraph breaks)
                blocks = [re.sub(r"[ \t]*\n[ \t]*", " ", b.strip()) for b in re.split(r"\n{2,}", t)]
                return ("\n\n".join(b for b in blocks if b)).strip()

            # Proper path: markdown -> HTML -> structured walk
            html = markdown.markdown(md_text)
            soup = BeautifulSoup(html, "html.parser")

            lines = []
            # Walk top-level blocks in order to preserve structure
            for node in soup.contents:
                name = getattr(node, "name", None)
                if not name:
                    # Stray text
                    text = " ".join(node.get_text(" ").split()).strip()
                    if text:
                        lines.append(text)
                    continue

                if name in ["h1", "h2", "h3", "h4", "h5", "h6"]:
                    txt = " ".join(node.get_text(" ").split()).strip()
                    if txt:
                        lines.append(txt)
                        lines.append("")  # blank line after heading

                elif name == "p":
                    txt = " ".join(node.get_text(" ").split()).strip()
                    if txt:
                        lines.append(txt)
                        lines.append("")

                elif name in ["ul", "ol"]:
                    # each li on its own line with bullet
                    for li in node.find_all("li", recursive=False):
                        li_text = " ".join(li.get_text(" ").split()).strip()
                        if li_text:
                            lines.append(f"• {li_text}")
                    lines.append("")

                elif name == "blockquote":
                    txt = " ".join(node.get_text(" ").split()).strip()
                    if txt:
                        lines.append(txt)
                        lines.append("")

                else:
                    # fallback for other blocks (code blocks, etc.)
                    txt = " ".join(node.get_text(" ").split()).strip()
                    if txt:
                        lines.append(txt)
                        lines.append("")

            # Remove excessive blank lines
            text = "\n".join(lines)
            text = re.sub(r"\n{3,}", "\n\n", text)

            # Tiny niceties: ensure no spaces before punctuation, collapse multiple spaces
            text = re.sub(r"\s+([,.;:!?])", r"\1", text)
            text = re.sub(r"[ \t]{2,}", " ", text)

            return text.strip()

        clean_text = _format_markdown_to_plain(raw_md)

        # Insert and lock
        textbox.insert("1.0", clean_text)
        textbox.configure(state="disabled")

    


    # =========================
    # Template Management System
    # =========================
    
    def _truncate_with_ellipsis(self, text: str, max_chars: int) -> str:
        """
        Truncate text to fit within max_chars, adding ellipsis if truncated.
        
        Args:
            text: The text to potentially truncate
            max_chars: Maximum number of characters allowed
            
        Returns:
            The original text if within limit, otherwise truncated with ellipsis
        """
        if len(text) <= max_chars:
            return text
        if max_chars <= 1:
            return "…"
        return text[: max_chars - 1] + "…"
    
    def _on_manual_only_toggled(self):
        idx = self._get_idx()
        if idx is None:
            return
        self.templates[idx]["manual_only"] = bool(self.manual_only_var.get())
        self._save_settings()
        self._refresh_template_list()


    def _refresh_template_list(self):
        """
        Refresh the template listbox display with current templates.
        
        Shows template names with asterisks for unsaved changes,
        and maintains current selection.
        """
        self.template_listbox.delete(0, tk.END)
        max_chars = self.listbox_char_width
        for idx, t in enumerate(self.templates):
            name = t.get("name", f"Template {idx+1}")
            manual_mark = "✓ " if t.get("manual_only", False) else "  "
            name_display = manual_mark + name
            if idx in self.unsaved_buffers:
                name_display += " *"
            self.template_listbox.insert(tk.END, self._truncate_with_ellipsis(name_display, max_chars))
        if self.current_index is not None and 0 <= self.current_index < len(self.templates):
            self.template_listbox.selection_clear(0, tk.END)
            self.template_listbox.selection_set(self.current_index)
            self.template_listbox.see(self.current_index)

    def _on_template_select(self, event=None):
        """
        Handle template selection from the listbox.
        
        Switches to the selected template in the editor.
        """
        sel = self.template_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if 0 <= idx < len(self.templates):
            self._select_template(idx)

    def _migrate_templates(self, raw):
        out = []
        if not raw:
            return out

        def _fix(t, i):
            # Normalize to dict
            if not isinstance(t, dict):
                return {"id": str(uuid.uuid4()), "name": f"Template {i+1}", "text": str(t or ""), "manual_only": False}

            return {
                "id": t.get("id") or str(uuid.uuid4()),
                "name": t.get("name") or f"Template {i+1}",
                "text": t.get("text") or "",
                "manual_only": bool(t.get("manual_only", False)),
            }

        if isinstance(raw, list):
            for i, t in enumerate(raw):
                out.append(_fix(t, i))
        return out


    def _html_to_text(self, html: str) -> str:
        """
        Convert HTML content to plain text.
        
        Removes HTML tags and converts br tags to newlines.
        
        Args:
            html: HTML string to convert
            
        Returns:
            Plain text version of the HTML
        """
        s = re.sub(r"(?i)<br\s*/?>", "\n", html)
        s = re.sub(r"(?is)<style.*?>.*?</style>", "", s)
        s = re.sub(r"(?is)<script.*?>.*?</script>", "", s)
        s = re.sub(r"(?s)<[^>]+>", "", s)
        return s.strip()

    def _new_template(self):
        """
        Create a new empty template.
        
        Prompts for template name and adds it to the template list.
        """
        name = simpledialog.askstring("New Template", "Template name:")
        if name is None:
            return
        name = name.strip() or f"Untitled {len(self.templates)+1}"
        self.templates.append({"id": str(uuid.uuid4()), "name": name, "text": "", "manual_only": False})
        self._save_settings()
        self._refresh_template_list()
        self._select_template(len(self.templates) - 1)
        self.toast.show(f"New template \"{name}\" created.", kind="success")

    def _rename_template(self):
        """
        Rename the currently selected template.
        
        Prompts for new name and updates the template.
        """
        idx = self._get_idx()
        if idx is None:
            self.toast.show("Select a template to rename.", kind="warning")
            return
        current = self.templates[idx]["name"]
        name = simpledialog.askstring("Rename Template", "New name:", initialvalue=current)
        if name is None:
            return
        name = name.strip() or current
        self.templates[idx]["name"] = name
        self._save_settings()
        self._update_editor_header()
        self.toast.show(f"Renamed to \"{name}\".", kind="success")

    def _move_up(self):
        """
        Move the currently selected template up in the list.
        
        Swaps position with the template above it.
        """
        idx = self._get_idx()
        if idx is None or idx == 0:
            return
        self.templates[idx-1], self.templates[idx] = self.templates[idx], self.templates[idx-1]
        self._swap_unsaved(idx, idx-1)
        self._save_settings()
        self._refresh_template_list()
        self._select_template(idx-1)
        self.toast.show("Moved up.", kind="info")

    def _move_down(self):
        """
        Move the currently selected template down in the list.
        
        Swaps position with the template below it.
        """
        idx = self._get_idx()
        if idx is None or idx >= len(self.templates) - 1:
            return
        self.templates[idx+1], self.templates[idx] = self.templates[idx], self.templates[idx+1]
        self._swap_unsaved(idx, idx+1)
        self._save_settings()
        self._refresh_template_list()
        self._select_template(idx+1)
        self.toast.show("Moved down.", kind="info")

    def _swap_unsaved(self, a, b):
        """
        Swap unsaved buffer entries when templates are reordered.
        
        Maintains unsaved changes when templates change positions.
        
        Args:
            a: First template index
            b: Second template index
        """
        a_has, b_has = a in self.unsaved_buffers, b in self.unsaved_buffers
        if a_has and b_has:
            self.unsaved_buffers[a], self.unsaved_buffers[b] = self.unsaved_buffers[b], self.unsaved_buffers[a]
        elif a_has and not b_has:
            self.unsaved_buffers[b] = self.unsaved_buffers[a]; del self.unsaved_buffers[a]
        elif b_has and not a_has:
            self.unsaved_buffers[a] = self.unsaved_buffers[b]; del self.unsaved_buffers[b]

    def _remove_template(self):
        """
        Remove the currently selected template.
        
        Deletes the template and updates unsaved buffer indices.
        """
        idx = self._get_idx()
        if idx is None:
            self.toast.show("Select a template to remove.", kind="warning")
            return
        name = self.templates[idx]["name"]
        del self.templates[idx]
        self._prune_recipient_overrides()
        new_unsaved = {}
        for k, v in self.unsaved_buffers.items():
            if k < idx: new_unsaved[k] = v
            elif k > idx: new_unsaved[k - 1] = v
        self.unsaved_buffers = new_unsaved
        self._save_settings()
        self._refresh_template_list()
        self.editor.delete("1.0", "end")
        self.current_index = None
        self.manual_only_var.set(False)
        self._update_editor_header()
        self.toast.show(f"Removed \"{name}\".", kind="warning")

    def _import_txts(self):
        """
        Import templates from .txt files.
        
        Allows bulk import of multiple text files as templates.
        """
        files = filedialog.askopenfilenames(title="Import .txt Templates", filetypes=[("Text Files", "*.txt")])
        if not files:
            return
        added = 0
        for fp in files:
            try:
                with open(fp, "r", encoding="utf-8") as f:
                    txt = f.read().rstrip()
                self.templates.append({
                    "id": str(uuid.uuid4()),
                    "name": Path(fp).stem,
                    "text": txt,
                    "manual_only": False
                })

                added += 1
            except Exception as e:
                self.toast.show(f"Failed to import {Path(fp).name}: {e}", kind="error")
        if added:
            self._save_settings()
            self._refresh_template_list()
            self._select_template(len(self.templates) - 1)
            self.toast.show(f"Imported {added} template(s).", kind="success")

    def _save_template_from_editor(self):
        """
        Save the current template from the editor.
        
        Commits editor content to the template and clears unsaved status.
        """
        idx = self._get_idx()
        if idx is None:
            self.toast.show("Select or create a template first.", kind="warning")
            return
        text = self.editor.get("1.0", "end").rstrip()
        self.templates[idx]["text"] = text
        if idx in self.unsaved_buffers:
            del self.unsaved_buffers[idx]
        self._save_settings()
        self._update_editor_header()
        self.toast.show(f"Saved \"{self.templates[idx]['name']}\".", kind="success")

    def _revert_current(self):
        """
        Revert the current template to its last saved state.
        
        Discards any unsaved changes in the editor.
        """
        idx = self._get_idx()
        if idx is None:
            return
        self._editor_set_text(self.templates[idx]["text"])
        if idx in self.unsaved_buffers:
            del self.unsaved_buffers[idx]
        self._update_editor_header()
        self.toast.show("Reverted to last saved.", kind="info")

    def _select_template(self, idx: int):
        """
        Select and load a template into the editor.
        
        Args:
            idx: Index of the template to select
        """
        self.current_index = idx
        self.manual_only_var.set(bool(self.templates[idx].get("manual_only", False)))
        text = self.unsaved_buffers.get(idx, self.templates[idx]["text"])
        self._editor_set_text(text)
        self._update_editor_header()
        self.template_listbox.selection_clear(0, tk.END)
        self.template_listbox.selection_set(idx)
        self.template_listbox.see(idx)

    def _editor_set_text(self, text: str):
        """
        Set the text content of the editor.
        
        Args:
            text: Text content to set in the editor
        """
        self.editor.configure(state="normal")
        self.editor.delete("1.0", "end")
        self.editor.insert("1.0", text)
        self.editor.edit_modified(False)

    def _update_editor_header(self):
        """
        Update the editor header to show current template name and unsaved status.
        
        Updates both the title label and unsaved indicator.
        """
        name = ""
        if self.current_index is not None and 0 <= self.current_index < len(self.templates):
            name = self.templates[self.current_index]["name"]
        self.editor_title.configure(text=f"Template Editor — {name}" if name else "Template Editor")
        self.unsaved_label.configure(text="● Unsaved" if self.current_index in self.unsaved_buffers else "")
        self._refresh_template_list()

    def _get_idx(self):
        """
        Get the current template index.
        
        Returns:
            Current template index or None if no template selected
        """
        return self.current_index

    # =========================
    # File pickers & CLEAR actions
    # =========================
    
    def _browse_csv(self):
        """
        Open file dialog to select a CSV file.
        
        Updates the CSV path and saves settings.
        """
        fn = filedialog.askopenfilename(title="Select CSV", filetypes=[("CSV Files", "*.csv")])
        if fn:
            self.csv_path.set(fn)
            self._save_settings()
            self.toast.show("CSV selected.", kind="info")

    def _clear_csv_path(self):
        """
        Clear the CSV file path.
        
        Removes the current CSV selection and saves settings.
        """
        self.csv_path.set("")
        self._save_settings()
        self.toast.show("Cleared CSV file.", kind="info")

    def _browse_resume(self):
        """
        Open file dialog to select a resume PDF.
        
        Updates the resume path and saves settings.
        """
        fn = filedialog.askopenfilename(title="Select Resume (PDF)", filetypes=[("PDF Files", "*.pdf")])
        if fn:
            self.resume_path.set(fn)
            self._save_settings()
            self.toast.show("Resume selected.", kind="info")

    def _clear_resume_path(self):
        """
        Clear the resume file path.
        
        Removes the current resume selection and saves settings.
        """
        self.resume_path.set("")
        self._save_settings()
        self.toast.show("Cleared resume.", kind="info")

    # =========================
    # Utilities
    # =========================
    
    def _on_editor_modified(self, event=None):
        """
        Handle editor content modification events.
        
        Tracks unsaved changes and updates the UI accordingly.
        """
        if not self.editor.edit_modified():
            return
        idx = self.current_index
        if idx is None:
            self.editor.edit_modified(False)
            return
        text = self.editor.get("1.0", "end").rstrip()
        saved = self.templates[idx]["text"]
        if text == saved:
            if idx in self.unsaved_buffers:
                del self.unsaved_buffers[idx]
        else:
            self.unsaved_buffers[idx] = text
        self._update_editor_header()
        self.editor.edit_modified(False)

    def _kb_save(self, event=None):
        """
        Handle keyboard shortcut for saving (Cmd+S).
        
        Returns:
            "break" to prevent default handling
        """
        if not getattr(self, "_licensed", False):
            self.toast.show("Enter a valid license key first.", kind="warning")
            return "break"
        self._save_template_from_editor()
        return "break"

    def _kb_rename(self, event=None):
        """
        Handle keyboard shortcut for renaming (Cmd+R).
        
        Returns:
            "break" to prevent default handling
        """
        if not getattr(self, "_licensed", False):
            self.toast.show("Enter a valid license key first.", kind="warning")
            return "break"
        self._rename_template()
        return "break"

    def _clear_all(self):
        """
        Clear all fields and data for the current profile.
        
        Resets data sources, templates, and other settings to defaults.
        """
        self.csv_path.set("")
        self.resume_path.set("")
        self.subject_template.set(DEFAULT_SUBJECT)
        self.gsheet_url.set("")
        self.data_source.set("sheet")
        self._gsheet_rows_cache = None
        self._last_headers_lower = []
        self.templates = []
        self.unsaved_buffers = {}
        self.current_index = None
        self._apply_source_enabled_state()
        self._refresh_template_list()
        self.editor.delete("1.0", "end")
        self._update_editor_header()
        self._save_settings()
        self.toast.show("Cleared fields for current profile.", kind="info")

    def _parse_name(self, full_name: str):
        """
        Parse a full name into first and last name components.
        
        Handles formats like "First Last" and "Last, First".
        
        Args:
            full_name: Full name string to parse
            
        Returns:
            Tuple of (first_name, last_name)
        """
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

    def _parse_bool(self, val: str) -> bool:
        """
        Parse a string value as a boolean.
        
        Args:
            val: String value to parse
            
        Returns:
            True if val represents a truthy value, False otherwise
        """
        return str(val).strip().lower() in {"1", "true", "yes", "y"}

    def _is_generate_true(self, row: dict) -> bool:
        """
        Check if a row should be processed for email generation.
        
        Args:
            row: Data row dictionary
            
        Returns:
            True if the row's 'generate' field indicates it should be processed
        """
        return self._parse_bool(row.get("generate", ""))

    def _is_email_valid(self, email: str) -> bool:
        """
        Validate an email address format.
        
        Args:
            email: Email address to validate
            
        Returns:
            True if email format is valid, False otherwise
        """
        return bool(email and EMAIL_RE.match(email))

    def _has_email_populated(self, row: dict) -> bool:
        """
        Check if a row has a valid email address.
        
        Args:
            row: Data row dictionary
            
        Returns:
            True if row contains a valid email address
        """
        resolver = self._make_resolver(row)
        email = resolver.get_email()
        return self._is_email_valid(email)


    def _replace_placeholders(self, text: str, placeholders: dict) -> str:
        """
        Replace placeholders in text with actual values.
        
        Args:
            text: Text containing placeholders like {Name}
            placeholders: Dictionary mapping placeholder names to values
            
        Returns:
            Text with placeholders replaced by actual values
        """
        out = text
        for key, value in placeholders.items():
            out = re.sub(r"\{" + re.escape(key) + r"\}", value, out, flags=re.IGNORECASE)
        return out

    # =========================
    # Resolver factory
    # =========================
    
    def _make_resolver(self, row: dict) -> PlaceholderResolver:
        """
        Create a PlaceholderResolver for a data row.
        
        Args:
            row: Data row dictionary
            
        Returns:
            PlaceholderResolver instance for the row
        """
        return PlaceholderResolver(self._last_headers_lower, row, self._parse_name)

    # =========================
    # Export templates to ZIP in Downloads
    # =========================
    
    def _safe_filename(self, name: str, max_len: int = 60) -> str:
        """
        Convert a template name to a safe filename.
        
        Removes invalid characters and truncates if necessary.
        
        Args:
            name: Template name to convert
            max_len: Maximum filename length
            
        Returns:
            Safe filename string
        """
        base = "".join(ch if (ch.isalnum() or ch in " -_.") else "_" for ch in (name or "template"))
        base = base.strip().strip(".")
        if not base:
            base = "template"
        if len(base) > max_len:
            base = base[:max_len]
        return base

    def _export_templates_zip(self):
        """
        Export all templates to a ZIP file in the Downloads folder.
        
        Creates a ZIP containing all templates as numbered text files.
        """
        if not self.templates:
            self.toast.show("No templates to export.", kind="warning")
            return

        downloads = Path.home() / "Downloads"
        dest_dir = downloads if downloads.exists() and downloads.is_dir() else Path.cwd()

        base_name = "My Email Templates"
        zip_path = dest_dir / f"{base_name}.zip"

        counter = 1
        while zip_path.exists():
            zip_path = dest_dir / f"{base_name} ({counter}).zip"
            counter += 1

        try:
            with zipfile.ZipFile(zip_path, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                for idx, tpl in enumerate(self.templates):
                    text = self.unsaved_buffers.get(idx, tpl.get("text", "")) or ""
                    name = tpl.get("name", f"Template {idx+1}")
                    fname = self._safe_filename(name)
                    arcname = f"{idx+1:02d}_{fname}.txt"
                    zf.writestr(arcname, text)

            self.toast.show(f"Exported to {zip_path.name}", kind="success")

            if IS_MAC:
                try:
                    subprocess.Popen(["open", "-R", str(zip_path)])
                except Exception:
                    pass

        except Exception as e:
            self.toast.show(f"Export failed: {e}", kind="error")

    # =========================
    # Generate
    # =========================
    
    def _handle_generate_clicked(self):
        """
        Handle the "Generate Emails" button click.
        
        Performs license validation and data loading before generating emails.
        """
        # 1) Quick local check + online re-check if >20 min
        if not self._check_license_smart():
            self.toast.show("License invalid or expired. Please validate.", kind="error")
            return

        # 2) Make sure the latest Sheet/CSV data is loaded
        if not self._ensure_latest_data():
            return

        # 3) Proceed with generation
        self._generate_emails()

    def _generate_emails(self):
        """
        Generate Outlook email drafts from templates and data.

        Validates inputs and calls engine to create drafts.
        """
        try:
            resume_path = self.resume_path.get().strip()
            subj_tpl = self.subject_template.get().strip()

            if not resume_path or not os.path.isfile(resume_path):
                self._flash_entry(self.resume_entry)
                self.toast.show("Please choose a valid resume PDF.", kind="error")
                return

            rows = self._read_eligible_rows()
            if not rows:
                self.toast.show("No rows to process (Generate != True or missing/invalid Email).", kind="warning")
                return

            if "{" not in subj_tpl:
                self.toast.show("Subject template has no placeholders. Continuing anyway.", kind="warning")

            # Generate via engine (engine handles preview rows and mappings internally)
            created = generate_emails(
                rows=rows,
                headers_lower=self._last_headers_lower,
                templates=self.templates,
                recipient_template_overrides=self.recipient_template_overrides,
                parse_name_fn=self._parse_name,
                is_generate_true_fn=self._is_generate_true,
                is_email_valid_fn=self._is_email_valid,
                subject_template=subj_tpl,
                resume_path=resume_path,
            )

            self.toast.show(f"Created {created} Outlook draft(s).", kind="success")
        except Exception as e:
            self.toast.show(f"Error: {e}", kind="error")

    def _flash_entry(self, entry: ctk.CTkEntry):
        """
        Flash an entry widget's border color to indicate an error.
        
        Args:
            entry: The entry widget to flash
        """
        orig = entry.cget("border_color")
        try:
            entry.configure(border_color="#ef4444")
            self.after(1600, lambda: entry.configure(border_color=orig))
        except Exception:
            pass

    # =========================
    # Lifecycle
    # =========================
    
    def _on_close(self):
        """
        Handle application close event.
        
        Prompts to save unsaved changes before closing.
        """
        if self._unsaved_changes_present():
            resp = messagebox.askyesnocancel(
                title="Unsaved changes",
                message=f"You have unsaved changes in \"{self.active_profile.get()}\". Save before quitting?",
                icon="warning"
            )
            if resp is None:
                return
            if resp:
                self._commit_all_unsaved_buffers()
                self.profiles[self.active_profile.get()] = self._collect_current_profile_state()
        self._save_settings()
        self.destroy()

    def _update_license_badge(self):
        """
        Update the license status badge in the UI.
        
        Shows green checkmark for licensed, red dot for unlicensed.
        """
        licensed = bool(getattr(self, "_licensed", False))
        txt = "✓ Licensed" if licensed else "● Unlicensed"
        color = "#10b981" if licensed else "#f87171"

        # Top-bar pill
        try:
            if hasattr(self, "license_status_top") and self.license_status_top is not None:
                self.license_status_top.configure(text=txt, text_color=color)
        except Exception:
            pass

        # (Optional) if you kept the old in-grid label, we'll update it too
        try:
            if hasattr(self, "license_status_lbl") and self.license_status_lbl is not None:
                self.license_status_lbl.configure(text=txt, text_color=color)
        except Exception:
            pass

    def _validate_license(self):
        """
        Validate the entered license key.
        
        Called by the Validate button to check license validity.
        """
        # Called by the Validate button
        if DEV_SKIP_LICENSE:
            self._licensed = True
            self._update_license_badge()
            self._apply_license_gate()
            self.toast.show("DEV: license bypassed.", kind="warning")
            return

        key = (self.license_key.get() or "").strip()
        if not key:
            self._licensed = False
            self._update_license_badge()
            self._apply_license_gate()
            self.toast.show("Please paste a license key first.", kind="warning")
            self.after(50, self.license_entry.focus_set)
            return

        ok, msg = self.license_mgr.validate_and_bind(key)  # <-- validate & auto-bind
        self._licensed = bool(ok)
        self._update_license_badge()
        self._apply_license_gate()

        if ok:
            self.toast.show(msg, kind="success")
            self._save_settings()
        else:
            self.toast.show(f"License invalid: {msg}", kind="error")
            # keep the UI locked

    def _apply_license_gate(self):
        """
        Apply license-based UI restrictions.
        
        Disables most functionality when unlicensed, enables when licensed.
        """
        licensed = getattr(self, "_licensed", False)

        # Always allow opening the modal
        try:
            self._set_enabled(self.license_btn, True)
        except Exception:
            pass

        widgets_to_gate = [
            self.source_seg, self.gsheet_entry, self.gsheet_load_btn, self.gsheet_clear_btn,
            self.gsheet_link_btn, self.csv_entry, self.csv_browse_btn, self.csv_clear_btn,
            self.resume_entry, self.resume_browse_btn, self.resume_clear_btn,
            getattr(self, "subj_entry", None),
            self.template_listbox,
            self.editor,
            getattr(self, "btn_preview", None),
            getattr(self, "btn_generate", None),
            getattr(self, "btn_clear_all", None),
            getattr(self, "btn_save_settings", None),
        ]

        for w in widgets_to_gate:
            if w is None:
                continue
            try:
                if w is self.editor:
                    w.configure(state="normal" if licensed else "disabled")
                else:
                    self._set_enabled(w, licensed)
            except Exception:
                pass

        if licensed:
            self._apply_source_enabled_state()


if __name__ == "__main__":
    app = EmailApp()
    app.mainloop()
