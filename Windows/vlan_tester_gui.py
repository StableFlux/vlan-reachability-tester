#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VLAN Reachability Tester — Windows 11 GUI Edition
Run: python vlan_tester_gui.py

Frozen/Store installs store config and results in
%LOCALAPPDATA%\\VLANReachabilityTester\\ (the app's install dir is read-only).
Running from source keeps them next to the script.
"""

import subprocess
import socket
import time
import os
import sys
import json
import shutil
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

import base64
import io

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
)

from _logo import LOGO_B64

# ─── File paths ───────────────────────────────────────────────────────────────

APP_DIR_NAME = "VLANReachabilityTester"


def _script_dir():
    """Directory of the exe (frozen) or the source script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def _user_data_dir():
    """Writable per-user data folder.

    Frozen (Store/MSIX install): %LOCALAPPDATA%\\VLANReachabilityTester\\ —
    the install dir under WindowsApps is read-only so we must write elsewhere.
    Source checkout: the script directory, so dev workflow is unchanged.
    """
    if not getattr(sys, "frozen", False):
        return _script_dir()
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser(r"~\AppData\Local")
    path = os.path.join(base, APP_DIR_NAME)
    try:
        os.makedirs(path, exist_ok=True)
    except Exception as exc:
        sys.stderr.write(f"[vlan-tester] Could not create {path}: {exc}\n")
    return path


SCRIPT_DIR   = _script_dir()
DATA_DIR     = _user_data_dir()
CONFIG_FILE  = os.path.join(DATA_DIR, "vlan_config.json")
RESULTS_FILE = os.path.join(DATA_DIR, "vlan_results.json")


def _migrate_legacy_file(filename):
    """One-time copy of a legacy file from next-to-the-exe to the user data dir.

    No-op when the data dir already has the file, or when we're running from
    source (SCRIPT_DIR == DATA_DIR)."""
    if SCRIPT_DIR == DATA_DIR:
        return
    new_path = os.path.join(DATA_DIR, filename)
    if os.path.exists(new_path):
        return
    old_path = os.path.join(SCRIPT_DIR, filename)
    if not os.path.exists(old_path):
        return
    try:
        shutil.copy2(old_path, new_path)
    except Exception as exc:
        sys.stderr.write(f"[vlan-tester] Could not migrate {filename}: {exc}\n")


_migrate_legacy_file("vlan_config.json")
_migrate_legacy_file("vlan_results.json")

_write_warnings_shown = set()


def _warn_write_failure(path, exc):
    """Report a write failure to stderr and raise a one-shot messagebox.

    Previously these failures were swallowed silently, which hid the v1.1.0
    Store-install persistence bug."""
    sys.stderr.write(f"[vlan-tester] Failed to write {path}: {exc}\n")
    if path in _write_warnings_shown:
        return
    _write_warnings_shown.add(path)
    try:
        messagebox.showerror(
            "Save failed",
            f"Could not write to:\n{path}\n\n{exc}\n\n"
            "Changes this session will not persist.",
        )
    except Exception:
        pass

DEFAULT_CONFIG = {
    "vlans": [],
    "ping_interval": 5,
    "ping_timeout":  2,
    "ping_count":    1,
    "selected_nic":  None,   # None = auto-detect; otherwise the IP of the chosen NIC
    "theme":         "system",   # "system" | "light" | "dark"
    "window_geometry": None,      # e.g. "1200x800+100+50" — restored on next launch
}


def format_subnet(subnet):
    """Display a stored prefix as a network address: '10.10.10.' -> '10.10.10.0'."""
    if not subnet:
        return subnet
    return subnet + "0" if subnet.endswith(".") else subnet


def parse_subnet(raw):
    """Accept '10.10.10.0' or '10.10.10.' from user input and normalise to the
    trailing-dot prefix form used internally for IP-startswith matching."""
    s = raw.strip()
    if not s:
        return s
    if s.endswith(".0"):
        return s[:-1]   # 10.10.10.0 -> 10.10.10.
    if not s.endswith("."):
        return s + "."
    return s

# ─── Theme Palettes ───────────────────────────────────────────────────────────
# Two palettes. A selection is resolved at startup from the saved config
# (default "system" reads the Windows app theme) and the chosen palette is
# blitted into the module globals (BG_DARK, CLR_GREEN, ...). Widgets look up
# these names at construction time, so changing theme requires an app restart
# — _cfg_apply offers to relaunch when the user picks a new theme.

_PALETTE_DARK = {
    "BG_DARK":    "#0d1117",
    "BG_PANEL":   "#161b22",
    "BG_HEADER":  "#1f2937",
    "CLR_GREEN":  "#22c55e",
    "CLR_RED":    "#ef4444",
    "CLR_YELLOW": "#fbbf24",
    "CLR_CYAN":   "#38bdf8",
    "CLR_PURPLE": "#a78bfa",
    "CLR_TEXT":   "#e2e8f0",
    "CLR_MUTED":  "#64748b",
    "CLR_BORDER": "#30363d",
    "CELL_GREEN": "#14532d",
    "CELL_RED":   "#7f1d1d",
    "CELL_GREY":  "#1e293b",
    "TV_SELECT":  "#1e3a5f",   # Treeview selection background
    "SELF_TAG":   "#86efac",   # "this is my VLAN" row highlight
}

# Mirrors the PDF export palette so screen and print look consistent.
_PALETTE_LIGHT = {
    "BG_DARK":    "#f8fafc",
    "BG_PANEL":   "#ffffff",
    "BG_HEADER":  "#e2e8f0",
    "CLR_GREEN":  "#16a34a",
    "CLR_RED":    "#dc2626",
    "CLR_YELLOW": "#d97706",
    "CLR_CYAN":   "#0891b2",
    "CLR_PURPLE": "#7c3aed",
    "CLR_TEXT":   "#0f172a",
    "CLR_MUTED":  "#64748b",
    "CLR_BORDER": "#cbd5e1",
    "CELL_GREEN": "#dcfce7",
    "CELL_RED":   "#fee2e2",
    "CELL_GREY":  "#f1f5f9",
    "TV_SELECT":  "#bae6fd",
    "SELF_TAG":   "#15803d",
}


def _windows_uses_light_theme():
    """Return True/False for Windows app theme, or None if unknown."""
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            val, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return bool(val)
    except OSError:
        return None


def resolve_theme(setting):
    """Translate a config setting ('system'|'light'|'dark') to 'light' or 'dark'."""
    if setting == "light":
        return "light"
    if setting == "dark":
        return "dark"
    return "light" if _windows_uses_light_theme() else "dark"


_active_theme = "dark"


def apply_palette(theme_name):
    """Blit the chosen palette into this module's globals."""
    global _active_theme
    _active_theme = "light" if theme_name == "light" else "dark"
    palette = _PALETTE_LIGHT if _active_theme == "light" else _PALETTE_DARK
    globals().update(palette)


def apply_titlebar_theme(root):
    """Match the Win11 title bar to the active theme. No-op on non-Windows
    and on Windows builds that don't implement the DWM attribute."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        hwnd = int(root.wm_frame(), 16)
        value = ctypes.c_int(1 if _active_theme == "dark" else 0)
        # DWMWA_USE_IMMERSIVE_DARK_MODE is attribute 20 on Win10 20H1+ / Win11.
        # Earlier 20H1 insider builds (19041 range) used attribute 19 — fall
        # back if the first call fails.
        res = ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(value), ctypes.sizeof(value),
        )
        if res != 0:
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 19, ctypes.byref(value), ctypes.sizeof(value),
            )
    except Exception:
        pass


# Apply dark as a safe default at import time; main() re-applies from config.
apply_palette("dark")

# ─── Config I/O ───────────────────────────────────────────────────────────────

def load_config():
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, encoding="utf-8") as f:
                data = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                data.setdefault(k, v)
            return data
    except Exception:
        pass
    return dict(DEFAULT_CONFIG)


def save_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2)
    except Exception as exc:
        _warn_write_failure(CONFIG_FILE, exc)


def load_results():
    try:
        if os.path.exists(RESULTS_FILE):
            with open(RESULTS_FILE, encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_results(results):
    try:
        with open(RESULTS_FILE, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2)
    except Exception as exc:
        _warn_write_failure(RESULTS_FILE, exc)


def export_report_pdf(vlan_names, vlans, results, current_vlan, my_ip, filename):
    """Generate a branded PDF VLAN reachability status report."""
    # Brand palette (matches the app's dark theme accent colours)
    C_CYAN   = colors.HexColor("#0891b2")
    C_DARK   = colors.HexColor("#0f172a")
    C_PANEL  = colors.HexColor("#1e293b")
    C_GREEN  = colors.HexColor("#16a34a")
    C_RED    = colors.HexColor("#dc2626")
    C_MUTED  = colors.HexColor("#64748b")
    C_TEXT   = colors.HexColor("#0f172a")
    C_BORDER = colors.HexColor("#cbd5e1")
    C_STRIPE = colors.HexColor("#f8fafc")
    CELL_GREEN = colors.HexColor("#dcfce7")
    CELL_RED   = colors.HexColor("#fee2e2")
    CELL_GREY  = colors.HexColor("#f1f5f9")

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    doc = SimpleDocTemplate(
        filename, pagesize=A4,
        leftMargin=16*mm, rightMargin=16*mm,
        topMargin=16*mm, bottomMargin=18*mm,
        title="VLAN Reachability Report",
        author="VLAN Reachability Tester",
    )

    title_style = ParagraphStyle(
        "Title", fontSize=22, textColor=C_CYAN, spaceAfter=2,
        fontName="Helvetica-Bold", leading=26,
    )
    subtitle_style = ParagraphStyle(
        "Subtitle", fontSize=10, textColor=C_MUTED, spaceAfter=8,
        fontName="Helvetica",
    )
    section_style = ParagraphStyle(
        "Section", fontSize=11, textColor=C_CYAN,
        spaceAfter=6, spaceBefore=14, fontName="Helvetica-Bold",
    )

    story = []

    # ── Header ────────────────────────────────────────────────────────────────
    logo_img = Image(io.BytesIO(base64.b64decode(LOGO_B64)),
                     width=22*mm, height=22*mm * 112 / 128)
    header_text = [
        Paragraph("VLAN REACHABILITY TESTER", title_style),
        Paragraph(
            f"Network Reachability Report &nbsp;&middot;&nbsp; Generated {now}",
            subtitle_style,
        ),
    ]
    header = Table(
        [[logo_img, header_text]],
        colWidths=[26*mm, doc.width - 26*mm],
    )
    header.setStyle(TableStyle([
        ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING",(0,0), (-1,-1), 0),
        ("TOPPADDING",  (0,0), (-1,-1), 0),
        ("BOTTOMPADDING",(0,0),(-1,-1), 0),
    ]))
    story.append(header)

    rule = Table([[""]], colWidths=[doc.width])
    rule.setStyle(TableStyle([("LINEBELOW", (0,0), (-1,-1), 1.5, C_CYAN)]))
    story.append(Spacer(1, 4))
    story.append(rule)

    # ── Summary / Current state ───────────────────────────────────────────────
    n        = len(vlan_names)
    total    = n * n
    tested   = len(results)
    reach    = sum(1 for v in results.values() if v.get("last") is True)
    blocked  = sum(1 for v in results.values() if v.get("last") is False)
    untested = total - tested

    summary = [
        ["CURRENT STATE", "", "SUMMARY", ""],
        ["Source VLAN", current_vlan or "UNKNOWN", "VLANs configured", str(n)],
        ["Local IP",    my_ip,                    "Pairs tested",     f"{tested} / {total}"],
        ["",            "",                       "Reachable",        str(reach)],
        ["",            "",                       "Blocked",          str(blocked)],
        ["",            "",                       "Untested",         str(untested)],
    ]
    col = doc.width / 4
    summary_tbl = Table(summary, colWidths=[col*0.9, col*1.1, col*0.9, col*1.1])
    summary_tbl.setStyle(TableStyle([
        ("FONTNAME",       (0,0), (-1,0),  "Helvetica-Bold"),
        ("TEXTCOLOR",      (0,0), (-1,0),  C_CYAN),
        ("FONTSIZE",       (0,0), (-1,0),  10),
        ("BOTTOMPADDING",  (0,0), (-1,0),  5),
        ("LINEBELOW",      (0,0), (1,0),   0.5, C_BORDER),
        ("LINEBELOW",      (2,0), (3,0),   0.5, C_BORDER),
        ("FONTNAME",       (0,1), (0,-1),  "Helvetica"),
        ("FONTNAME",       (2,1), (2,-1),  "Helvetica"),
        ("TEXTCOLOR",      (0,1), (0,-1),  C_MUTED),
        ("TEXTCOLOR",      (2,1), (2,-1),  C_MUTED),
        ("FONTNAME",       (1,1), (1,-1),  "Helvetica-Bold"),
        ("FONTNAME",       (3,1), (3,-1),  "Helvetica-Bold"),
        ("TEXTCOLOR",      (1,1), (1,-1),  C_TEXT),
        ("TEXTCOLOR",      (3,1), (3,-1),  C_TEXT),
        ("FONTSIZE",       (0,1), (-1,-1), 9),
        ("VALIGN",         (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",     (0,1), (-1,-1), 3),
        ("BOTTOMPADDING",  (0,1), (-1,-1), 3),
        ("LEFTPADDING",    (0,0), (-1,-1), 2),
    ]))
    story.append(Spacer(1, 10))
    story.append(summary_tbl)

    # ── VLAN Definitions ──────────────────────────────────────────────────────
    story.append(Paragraph("VLAN DEFINITIONS", section_style))
    def_data = [["VLAN", "SUBNET", "TARGET IP", "DEVICE"]]
    for name in vlan_names:
        v = vlans[name]
        def_data.append([name, format_subnet(v["subnet"]), v["target"], v.get("label", "")])

    def_tbl = Table(def_data, colWidths=[col*0.7, col*0.9, col*0.9, col*1.5])
    def_tbl.setStyle(TableStyle([
        ("BACKGROUND",     (0,0), (-1,0),  C_PANEL),
        ("TEXTCOLOR",      (0,0), (-1,0),  colors.white),
        ("FONTNAME",       (0,0), (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",       (0,0), (-1,0),  9),
        ("FONTNAME",       (0,1), (-1,-1), "Helvetica"),
        ("FONTSIZE",       (0,1), (-1,-1), 9),
        ("TEXTCOLOR",      (0,1), (-1,-1), C_TEXT),
        ("FONTNAME",       (0,1), (0,-1),  "Helvetica-Bold"),
        ("TEXTCOLOR",      (0,1), (0,-1),  C_CYAN),
        ("GRID",           (0,0), (-1,-1), 0.3, C_BORDER),
        ("VALIGN",         (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",     (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",  (0,0), (-1,-1), 5),
        ("LEFTPADDING",    (0,0), (-1,-1), 8),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, C_STRIPE]),
    ]))
    story.append(def_tbl)

    # ── Reachability Matrix ───────────────────────────────────────────────────
    story.append(Paragraph("REACHABILITY MATRIX", section_style))
    matrix_data = [[""] + [nm[:7] for nm in vlan_names]]
    cell_colors = []
    for i, frm in enumerate(vlan_names):
        row = [frm]
        for j, to in enumerate(vlan_names):
            entry = results.get(f"{frm}->{to}")
            if entry is None:
                row.append("-")
                cell_colors.append(("BACKGROUND", (j+1, i+1), (j+1, i+1), CELL_GREY))
                cell_colors.append(("TEXTCOLOR",  (j+1, i+1), (j+1, i+1), C_MUTED))
            elif entry.get("last"):
                rtt = entry.get("rtt")
                row.append(f"{rtt:.0f}" if rtt is not None else "OK")
                cell_colors.append(("BACKGROUND", (j+1, i+1), (j+1, i+1), CELL_GREEN))
                cell_colors.append(("TEXTCOLOR",  (j+1, i+1), (j+1, i+1), C_GREEN))
            else:
                row.append("X")
                cell_colors.append(("BACKGROUND", (j+1, i+1), (j+1, i+1), CELL_RED))
                cell_colors.append(("TEXTCOLOR",  (j+1, i+1), (j+1, i+1), C_RED))
        matrix_data.append(row)

    avail = doc.width - 70
    cell_w = min(40, avail / max(n, 1))
    mx_widths = [70] + [cell_w] * n
    matrix_tbl = Table(matrix_data, colWidths=mx_widths, rowHeights=[22] + [22]*n)
    mx_style = [
        ("BACKGROUND",    (0,0), (-1,0),  C_PANEL),
        ("TEXTCOLOR",     (1,0), (-1,0),  colors.white),
        ("FONTNAME",      (0,0), (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,0),  8),
        ("ALIGN",         (0,0), (-1,-1), "CENTER"),
        ("FONTNAME",      (0,1), (0,-1),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,1), (0,-1),  8),
        ("TEXTCOLOR",     (0,1), (0,-1),  colors.white),
        ("BACKGROUND",    (0,1), (0,-1),  C_PANEL),
        ("FONTNAME",      (1,1), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE",      (1,1), (-1,-1), 8),
        ("GRID",          (0,0), (-1,-1), 0.4, colors.white),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
    ]
    mx_style.extend(cell_colors)
    matrix_tbl.setStyle(TableStyle(mx_style))
    story.append(matrix_tbl)

    # Legend
    legend = Table(
        [[u"\u25A0", "Reachable", u"\u25A0", "Blocked", u"\u25A0", "Not yet tested"]],
        colWidths=[10, 60, 10, 50, 10, 80],
    )
    legend.setStyle(TableStyle([
        ("TEXTCOLOR",  (0,0), (0,0), C_GREEN),
        ("TEXTCOLOR",  (2,0), (2,0), C_RED),
        ("TEXTCOLOR",  (4,0), (4,0), C_MUTED),
        ("TEXTCOLOR",  (1,0), (1,0), C_TEXT),
        ("TEXTCOLOR",  (3,0), (3,0), C_TEXT),
        ("TEXTCOLOR",  (5,0), (5,0), C_TEXT),
        ("FONTSIZE",   (0,0), (0,0), 13),
        ("FONTSIZE",   (2,0), (2,0), 13),
        ("FONTSIZE",   (4,0), (4,0), 13),
        ("FONTSIZE",   (1,0), (-1,-1), 9),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING",(0,0), (-1,-1), 2),
    ]))
    story.append(Spacer(1, 6))
    story.append(legend)

    # ── Detailed Results ──────────────────────────────────────────────────────
    story.append(Paragraph("DETAILED RESULTS", section_style))
    detail = [["SOURCE", "DESTINATION", "STATUS", "RTT", "LAST TESTED"]]
    status_styles = []
    for frm in vlan_names:
        for to in vlan_names:
            entry = results.get(f"{frm}->{to}")
            if entry is None:
                status, rtt_s, time_s, sc = "Untested", "--", "--", C_MUTED
            elif entry.get("last"):
                rtt = entry.get("rtt")
                rtt_s = f"{rtt:.1f} ms" if rtt is not None else "--"
                status, time_s, sc = "Reachable", entry.get("time", "--"), C_GREEN
            else:
                status, rtt_s, time_s, sc = "Blocked", "--", entry.get("time", "--"), C_RED
            detail.append([frm, to, status, rtt_s, time_s])
            r = len(detail) - 1
            status_styles.append(("TEXTCOLOR", (2, r), (2, r), sc))
            status_styles.append(("FONTNAME",  (2, r), (2, r), "Helvetica-Bold"))

    dt_widths = [col*0.7, col*0.85, col*0.8, col*0.7, col*0.95]
    detail_tbl = Table(detail, colWidths=dt_widths, repeatRows=1)
    dt_style = [
        ("BACKGROUND",     (0,0), (-1,0),  C_PANEL),
        ("TEXTCOLOR",      (0,0), (-1,0),  colors.white),
        ("FONTNAME",       (0,0), (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",       (0,0), (-1,0),  9),
        ("FONTNAME",       (0,1), (-1,-1), "Helvetica"),
        ("FONTSIZE",       (0,1), (-1,-1), 8.5),
        ("TEXTCOLOR",      (0,1), (-1,-1), C_TEXT),
        ("FONTNAME",       (0,1), (0,-1),  "Helvetica-Bold"),
        ("TEXTCOLOR",      (0,1), (0,-1),  C_CYAN),
        ("GRID",           (0,0), (-1,-1), 0.3, C_BORDER),
        ("VALIGN",         (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",     (0,0), (-1,-1), 4),
        ("BOTTOMPADDING",  (0,0), (-1,-1), 4),
        ("LEFTPADDING",    (0,0), (-1,-1), 8),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, C_STRIPE]),
    ]
    dt_style.extend(status_styles)
    detail_tbl.setStyle(TableStyle(dt_style))
    story.append(detail_tbl)

    # ── Footer on every page ──────────────────────────────────────────────────
    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(C_BORDER)
        canvas.setLineWidth(0.4)
        canvas.line(16*mm, 14*mm, doc.pagesize[0] - 16*mm, 14*mm)
        canvas.setFont("Helvetica-Bold", 8)
        canvas.setFillColor(C_CYAN)
        canvas.drawString(16*mm, 9*mm, "VLAN Reachability Tester")
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(C_MUTED)
        canvas.drawCentredString(doc.pagesize[0] / 2, 9*mm, now)
        canvas.drawRightString(doc.pagesize[0] - 16*mm, 9*mm, f"Page {doc.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)

# ─── Network Helpers ──────────────────────────────────────────────────────────

def get_network_interfaces():
    """Return list of {name, ip, alias} dicts for all active IPv4 adapters."""
    ifaces = []
    try:
        out = subprocess.check_output(
            ["ipconfig", "/all"], text=True, encoding="utf-8", errors="replace",
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        current_name = None
        for line in out.splitlines():
            stripped = line.rstrip()
            if stripped and not stripped.startswith(" ") and stripped.endswith(":"):
                current_name = stripped.rstrip(":").strip()
            elif current_name and "IPv4 Address" in stripped:
                raw = stripped.split(":")[-1].strip().replace("(Preferred)", "").strip()
                if raw and not raw.startswith("127.") and not raw.startswith("169.254."):
                    # Extract alias: "Wireless LAN adapter Wi-Fi" → "Wi-Fi"
                    alias = current_name
                    lower = current_name.lower()
                    idx = lower.find(" adapter ")
                    if idx != -1:
                        alias = current_name[idx + len(" adapter "):]
                    ifaces.append({"name": current_name, "ip": raw, "alias": alias})
    except Exception:
        pass
    return ifaces


def _wait_for_new_ip(alias, old_ip, retries=15, delay=1.0):
    """Poll get_network_interfaces() until the alias has a new non-APIPA IP."""
    for _ in range(retries):
        time.sleep(delay)
        for iface in get_network_interfaces():
            if iface["alias"] == alias:
                ip = iface["ip"]
                if ip and ip != old_ip and not ip.startswith("169.254."):
                    return ip
    return None


def get_local_ips():
    ips = set()
    try:
        import netifaces
        for iface in netifaces.interfaces():
            for addr in netifaces.ifaddresses(iface).get(netifaces.AF_INET, []):
                ip = addr.get("addr", "")
                if ip and not ip.startswith("127.") and not ip.startswith("169.254."):
                    ips.add(ip)
    except ImportError:
        pass

    if not ips:
        try:
            for ip in socket.gethostbyname_ex(socket.gethostname())[2]:
                if not ip.startswith("127.") and not ip.startswith("169.254."):
                    ips.add(ip)
        except Exception:
            pass

    if not ips:
        try:
            out = subprocess.check_output(
                ["ipconfig"], text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            for line in out.splitlines():
                if "IPv4 Address" in line:
                    ip = line.split(":")[-1].strip()
                    if not ip.startswith("127.") and not ip.startswith("169.254."):
                        ips.add(ip)
        except Exception:
            pass

    return list(ips)


def detect_vlan(ips, vlans):
    for ip in ips:
        for name, vlan in vlans.items():
            if ip.startswith(vlan["subnet"]):
                return name, ip
    return None, (ips[0] if ips else "unknown")


def ping(host, count=1, timeout=2, source_ip=None):
    cmd = ["ping", "-n", str(count), "-w", str(timeout * 1000)]
    if source_ip:
        cmd += ["-S", source_ip]
    cmd.append(host)
    try:
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if result.returncode == 0:
            for line in result.stdout.splitlines():
                if "Average" in line:
                    try:
                        rtt = float(line.split("=")[-1].strip().replace("ms", ""))
                        return True, rtt
                    except Exception:
                        pass
            return True, None
        return False, None
    except Exception:
        return False, None

# ─── VLAN Add / Edit Dialog ───────────────────────────────────────────────────

class VlanDialog(tk.Toplevel):
    """Modal dialog for adding or editing a VLAN entry."""

    def __init__(self, parent, title, initial=None):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=BG_HEADER)
        self.resizable(False, False)
        self.result = None

        initial = initial or {}

        pad = {"padx": 14, "pady": 6}

        fields = [
            ("VLAN Name",   "name",   initial.get("name",   "")),
            ("Subnet",      "subnet", format_subnet(initial.get("subnet", ""))),
            ("Target IP",   "target", initial.get("target", "")),
            ("Device Label","label",  initial.get("label",  "")),
        ]

        self.vars = {}
        for i, (lbl, key, val) in enumerate(fields):
            tk.Label(
                self, text=lbl, font=("Consolas", 10),
                fg=CLR_MUTED, bg=BG_HEADER, anchor="e", width=14,
            ).grid(row=i, column=0, **pad, sticky="e")

            var = tk.StringVar(value=val)
            entry = tk.Entry(
                self, textvariable=var, font=("Consolas", 11),
                bg=BG_DARK, fg=CLR_TEXT, insertbackground=CLR_TEXT,
                relief="flat", bd=6, width=26,
            )
            entry.grid(row=i, column=1, **pad, sticky="ew")
            self.vars[key] = var

        # Hint row
        tk.Label(
            self, text='Subnet e.g. "192.168.10.0"',
            font=("Consolas", 8), fg=CLR_MUTED, bg=BG_HEADER,
        ).grid(row=len(fields), column=0, columnspan=2, padx=14, pady=(0, 4))

        # Buttons
        btn_row = tk.Frame(self, bg=BG_HEADER)
        btn_row.grid(row=len(fields)+1, column=0, columnspan=2, pady=(4, 14))

        _btn(btn_row, "Save", CLR_GREEN, BG_DARK, self._on_save).pack(side="left", padx=8)
        _btn(btn_row, "Cancel", CLR_MUTED, BG_DARK, self.destroy).pack(side="left", padx=8)

        self.grab_set()
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self.destroy)

        # Centre on parent
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

    def _on_save(self):
        name   = self.vars["name"].get().strip().upper()
        subnet_raw = self.vars["subnet"].get().strip()
        target = self.vars["target"].get().strip()
        label  = self.vars["label"].get().strip()

        if not name or not subnet_raw or not target:
            messagebox.showwarning("Missing fields",
                                   "Name, Subnet, and Target IP are required.",
                                   parent=self)
            return

        subnet = parse_subnet(subnet_raw)

        self.result = {"name": name, "subnet": subnet, "target": target, "label": label}
        self.destroy()

# ─── Helpers ──────────────────────────────────────────────────────────────────

def _btn(parent, text, bg, fg, cmd, **kw):
    return tk.Button(
        parent, text=text, bg=bg, fg=fg, command=cmd,
        font=("Consolas", 10, "bold"),
        relief="flat", padx=14, pady=6, cursor="hand2", bd=0,
        **kw,
    )


def _section_label(parent, text):
    tk.Label(
        parent, text=text,
        font=("Consolas", 10, "bold"),
        fg=CLR_CYAN, bg=BG_DARK,
    ).pack(anchor="w", pady=(0, 4))

def _status_label(ok):
    if ok is True:
        return "REACHABLE"
    if ok is False:
        return "BLOCKED"
    return "UNTESTED"


class MatrixTooltip:
    """Hover tooltip for a reachability matrix cell. Shows the ping history
    stats we keep per (source, destination) pair."""

    SHOW_DELAY_MS = 400

    def __init__(self, app, canvas):
        self.app = app
        self.canvas = canvas
        self.tip_window = None
        self.after_id = None
        self.current_cell = None

        canvas.bind("<Motion>", self._on_motion)
        canvas.bind("<Leave>", self._on_leave)

    def _cell_from_xy(self, x, y):
        names = self.app.vlan_names
        if not names:
            return None
        lw, lh, px = self.app.LABEL_W, self.app.LABEL_H, self.app.CELL_PX
        if x < lw or y < lh:
            return None
        col = (x - lw) // px
        row = (y - lh) // px
        if 0 <= col < len(names) and 0 <= row < len(names):
            return (names[row], names[col])
        return None

    def _on_motion(self, event):
        cell = self._cell_from_xy(event.x, event.y)
        if cell == self.current_cell:
            return
        self._hide()
        self.current_cell = cell
        if cell is not None:
            x_root, y_root = event.x_root, event.y_root
            self.after_id = self.canvas.after(
                self.SHOW_DELAY_MS,
                lambda: self._show(x_root, y_root),
            )

    def _on_leave(self, _event):
        self.current_cell = None
        self._hide()

    def _show(self, x_root, y_root):
        if self.current_cell is None:
            return
        src, dst = self.current_cell
        text = self._build_text(src, dst)

        self.tip_window = tk.Toplevel(self.canvas)
        self.tip_window.overrideredirect(True)
        self.tip_window.attributes("-topmost", True)
        self.tip_window.configure(bg=CLR_BORDER)
        tk.Label(
            self.tip_window, text=text,
            font=("Consolas", 9),
            fg=CLR_TEXT, bg=BG_PANEL,
            justify="left", padx=10, pady=6,
        ).pack(padx=1, pady=1)
        self.tip_window.geometry(f"+{x_root + 14}+{y_root + 18}")

    def _hide(self):
        if self.after_id is not None:
            try:
                self.canvas.after_cancel(self.after_id)
            except Exception:
                pass
            self.after_id = None
        if self.tip_window is not None:
            try:
                self.tip_window.destroy()
            except Exception:
                pass
            self.tip_window = None

    def _build_text(self, src, dst):
        key = f"{src}->{dst}"
        with self.app._lock:
            entry = dict(self.app.results.get(key, {}))
        history = entry.get("history") or []

        lines = [f"{src}  →  {dst}"]
        if not history:
            if not entry:
                lines.append("No data yet")
                return "\n".join(lines)
            rtt = entry.get("rtt")
            rtt_s = f"{rtt:.1f} ms" if rtt is not None else "—"
            lines.append(f"Status:  {_status_label(entry.get('last'))}")
            lines.append(f"Last:    {rtt_s}  at {entry.get('time', '?')}")
            return "\n".join(lines)

        total = len(history)
        oks   = [h for h in history if h.get("ok") is True]
        rtts  = [h["rtt"] for h in history if h.get("ok") is True and h.get("rtt") is not None]
        rate  = int(round(100.0 * len(oks) / total))
        last  = history[-1]
        last_rtt = last.get("rtt")
        last_rtt_s = f"{last_rtt:.1f} ms" if last_rtt is not None else "—"

        lines.append(f"Status:  {_status_label(last.get('ok'))}")
        lines.append(f"Last:    {last_rtt_s}  at {last.get('time', '?')}")
        lines.append(f"Success: {len(oks)}/{total}  ({rate}%)")
        if rtts:
            avg = sum(rtts) / len(rtts)
            lines.append(f"RTT avg: {avg:.1f} ms   (min {min(rtts):.1f}, max {max(rtts):.1f})")
        return "\n".join(lines)


# ─── Main Application ─────────────────────────────────────────────────────────

class VlanTesterApp:
    CELL_PX = 54
    LABEL_W = 76
    LABEL_H = 28

    def __init__(self, root: tk.Tk):
        self.root    = root
        self.config  = load_config()
        self._migrate_nic_config()   # convert stored IP → alias if config is old format
        self.results = load_results()
        self._derive_vlans()

        self.sweep_count  = 0
        self.paused       = False
        self.running      = True
        self.current_vlan = None
        self.my_ip        = "detecting…"
        self.countdown    = 0
        self._lock        = threading.Lock()

        self._configure_window()
        self._apply_notebook_style()
        self._build_notebook()
        self._build_monitor_tab()
        self._build_config_tab()
        self._start_worker()
        self._tick()

    # ── Setup ─────────────────────────────────────────────────────────────────

    def _migrate_nic_config(self):
        """Ensure selected_nic holds an alias. Convert legacy IP format if possible.
        If no NIC is saved yet, default to the first available adapter."""
        val   = self.config.get("selected_nic")
        ifaces = get_network_interfaces()

        if val and all(c.isdigit() or c == "." for c in val):
            # Legacy IP stored — try to convert to alias
            match = next((i for i in ifaces if i["ip"] == val), None)
            if match:
                self.config["selected_nic"] = match["alias"]
                save_config(self.config)
            # If IP not found right now, leave as-is; worker handles it gracefully

        if not self.config.get("selected_nic") and ifaces:
            # Nothing selected yet — default to first adapter
            self.config["selected_nic"] = ifaces[0]["alias"]
            save_config(self.config)

    def _configure_window(self):
        self.root.title("VLAN Reachability Tester")
        self.root.configure(bg=BG_DARK)
        self.root.minsize(960, 640)

        saved = self.config.get("window_geometry")
        if saved and isinstance(saved, str):
            try:
                self.root.geometry(saved)
            except Exception:
                self.root.geometry("1200x800")
        else:
            self.root.geometry("1200x800")

        # Paint the Win11 title bar to match the active theme.
        self.root.update_idletasks()
        apply_titlebar_theme(self.root)

    def _snapshot_geometry(self):
        """Capture window geometry into config so it persists across launches."""
        try:
            self.config["window_geometry"] = self.root.geometry()
        except Exception:
            pass

    def _derive_vlans(self):
        self.vlans      = {}
        self.vlan_names = []
        for v in self.config.get("vlans", []):
            name = v["name"]
            self.vlans[name] = {
                "subnet": v["subnet"],
                "target": v["target"],
                "label":  v.get("label", ""),
            }
            self.vlan_names.append(name)

    def _apply_notebook_style(self):
        style = ttk.Style()
        style.theme_use("clam")

    # ── Notebook ──────────────────────────────────────────────────────────────

    def _build_notebook(self):
        # Global header
        hdr = tk.Frame(self.root, bg=BG_HEADER, height=58)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(
            hdr, text="⬡  VLAN REACHABILITY TESTER",
            font=("Consolas", 16, "bold"),
            fg=CLR_CYAN, bg=BG_HEADER,
        ).pack(side="left", padx=22, pady=12)
        self.ts_var = tk.StringVar()
        tk.Label(hdr, textvariable=self.ts_var,
                 font=("Consolas", 10), fg=CLR_MUTED, bg=BG_HEADER,
                 ).pack(side="right", padx=22)

        # Custom tab bar
        tab_bar = tk.Frame(self.root, bg=BG_PANEL, height=44)
        tab_bar.pack(fill="x")
        tab_bar.pack_propagate(False)

        # Tab buttons (left). Each tab is a button stacked above a 3-px
        # underline strip; the strip lights up in the accent colour for the
        # active tab, giving an unambiguous selection cue in both themes.
        self._tab_btns = {}
        self._tab_underlines = {}
        for name, label in [("monitor", "  ▶  Monitor  "), ("config", "  ⚙  Config  ")]:
            wrap = tk.Frame(tab_bar, bg=BG_PANEL)
            wrap.pack(side="left", fill="y")
            btn = tk.Button(
                wrap, text=label,
                font=("Consolas", 11, "bold"),
                relief="flat", bd=0, cursor="hand2",
                padx=0, pady=0,
                command=lambda n=name: self._show_tab(n),
            )
            btn.pack(side="top", fill="both", expand=True)
            underline = tk.Frame(wrap, bg=BG_PANEL, height=3)
            underline.pack(side="bottom", fill="x")
            self._tab_btns[name] = btn
            self._tab_underlines[name] = underline

        # Monitor action buttons (right) — hidden when Config is active
        bkw = dict(font=("Consolas", 10, "bold"), relief="flat",
                   padx=14, pady=5, cursor="hand2", bd=0)
        self._monitor_btn_frame = tk.Frame(tab_bar, bg=BG_PANEL)
        self._monitor_btn_frame.pack(side="right", fill="y", padx=(0, 8))

        tk.Button(self._monitor_btn_frame, text="🗑  CLEAR MATRIX",
                  bg=BG_PANEL, fg=CLR_RED,
                  command=self._clear_matrix, **bkw).pack(side="right", pady=7, padx=(4, 0))
        tk.Button(self._monitor_btn_frame, text="💾  REPORT",
                  bg=CLR_CYAN, fg=BG_DARK,
                  command=self._do_export, **bkw).pack(side="right", pady=7, padx=(4, 0))
        tk.Button(self._monitor_btn_frame, text="🔄  RENEW IP",
                  bg=CLR_PURPLE, fg=BG_DARK,
                  command=self._renew_ip, **bkw).pack(side="right", pady=7, padx=(4, 0))
        self.pause_btn = tk.Button(self._monitor_btn_frame, text="⏸  RUNNING",
                                   bg=CLR_GREEN, fg=BG_DARK,
                                   command=self._toggle_pause, **bkw)
        self.pause_btn.pack(side="right", pady=7, padx=(4, 0))

        # Config action button (right) — hidden when Monitor is active
        self._config_btn_frame = tk.Frame(tab_bar, bg=BG_PANEL)
        self._config_btn_frame.pack(side="right", fill="y", padx=(0, 8))
        self.apply_msg = tk.StringVar()
        tk.Label(self._config_btn_frame, textvariable=self.apply_msg,
                 font=("Consolas", 10), fg=CLR_GREEN, bg=BG_PANEL,
                 ).pack(side="left", padx=(0, 10))
        tk.Button(self._config_btn_frame, text="✔  Apply & Restart Sweep",
                  bg=CLR_GREEN, fg=BG_DARK, command=self._cfg_apply, **bkw,
                  ).pack(side="right", pady=7)
        tk.Button(self._config_btn_frame, text="⬆  EXPORT",
                  bg=CLR_CYAN, fg=BG_DARK, command=self._cfg_export_config,
                  **bkw).pack(side="right", pady=7, padx=(0, 6))
        tk.Button(self._config_btn_frame, text="⬇  IMPORT",
                  bg=CLR_PURPLE, fg=BG_DARK, command=self._cfg_import_config,
                  **bkw).pack(side="right", pady=7, padx=(0, 6))

        # Content frames
        self.monitor_tab = tk.Frame(self.root, bg=BG_DARK)
        self.config_tab  = tk.Frame(self.root, bg=BG_DARK)

        self._active_tab = None
        self._show_tab("monitor")

    def _show_tab(self, name):
        if self._active_tab == "monitor":
            self.monitor_tab.pack_forget()
        elif self._active_tab == "config":
            self.config_tab.pack_forget()

        self._active_tab = name

        if name == "monitor":
            self.monitor_tab.pack(fill="both", expand=True)
            self._config_btn_frame.pack_forget()
            self._monitor_btn_frame.pack(side="right", fill="y", padx=(0, 8))
        else:
            self.config_tab.pack(fill="both", expand=True)
            self._monitor_btn_frame.pack_forget()
            self._config_btn_frame.pack(side="right", fill="y", padx=(0, 8))

        for n, btn in self._tab_btns.items():
            is_active = n == name
            btn.config(
                bg=BG_HEADER if is_active else BG_PANEL,
                fg=CLR_CYAN  if is_active else CLR_MUTED,
            )
            self._tab_underlines[n].config(bg=CLR_CYAN if is_active else BG_PANEL)

    # ── Monitor Tab ───────────────────────────────────────────────────────────

    def _build_monitor_tab(self):
        for w in self.monitor_tab.winfo_children():
            w.destroy()

        self._build_statusbar(self.monitor_tab)

        self.monitor_content = tk.Frame(self.monitor_tab, bg=BG_DARK)
        self.monitor_content.pack(fill="both", expand=True, padx=12, pady=8)

        if not self.vlan_names:
            tk.Label(
                self.monitor_content,
                text="No VLANs configured.\n\nGo to the Config tab to add your VLANs.",
                font=("Consolas", 13), fg=CLR_MUTED, bg=BG_DARK,
            ).pack(expand=True)
        else:
            left = tk.Frame(self.monitor_content, bg=BG_DARK)
            left.pack(side="left", fill="both", expand=True, padx=(0, 6))
            _section_label(left, "CURRENT VLAN SWEEP")
            self._build_sweep_table(left)

            right = tk.Frame(self.monitor_content, bg=BG_DARK)
            right.pack(side="right", fill="both", padx=(6, 0))
            _section_label(right, "FULL REACHABILITY MATRIX  (row = source, col = dest)")
            self._build_matrix(right)

        self._build_footer(self.monitor_tab)

    def _build_statusbar(self, parent):
        bar = tk.Frame(parent, bg=BG_PANEL, height=44)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        tk.Label(bar, text="", bg=BG_PANEL, width=2).pack(side="left")

        def pill(label, var, fg):
            tk.Label(bar, text=label, font=("Consolas", 10),
                     fg=CLR_MUTED, bg=BG_PANEL).pack(side="left")
            tk.Label(bar, textvariable=var, font=("Consolas", 11, "bold"),
                     fg=fg, bg=BG_PANEL).pack(side="left")

        self.vlan_var  = tk.StringVar(value="Detecting…")
        self.ip_var    = tk.StringVar(value="…")
        self.nic_var   = tk.StringVar(value="auto")
        self.sweep_var = tk.StringVar(value="0")

        pill("VLAN ", self.vlan_var, CLR_CYAN)
        tk.Label(bar, text="  │  ", font=("Consolas", 10),
                 fg=CLR_BORDER, bg=BG_PANEL).pack(side="left")
        pill("IP ", self.ip_var, CLR_TEXT)
        tk.Label(bar, text="  │  ", font=("Consolas", 10),
                 fg=CLR_BORDER, bg=BG_PANEL).pack(side="left")
        pill("NIC ", self.nic_var, CLR_YELLOW)
        tk.Label(bar, text="  │  ", font=("Consolas", 10),
                 fg=CLR_BORDER, bg=BG_PANEL).pack(side="left")
        pill("Sweep ", self.sweep_var, CLR_TEXT)

        self.countdown_var = tk.StringVar()
        tk.Label(bar, textvariable=self.countdown_var,
                 font=("Consolas", 10), fg=CLR_MUTED, bg=BG_PANEL,
                 ).pack(side="right", padx=(0, 14))

        self.status_lbl = tk.Label(
            bar, text="● ACTIVE",
            font=("Consolas", 11, "bold"), fg=CLR_GREEN, bg=BG_PANEL,
        )
        self.status_lbl.pack(side="right", padx=(0, 18))

    def _build_sweep_table(self, parent):
        style = ttk.Style()
        style.configure("V.Treeview",
                        background=BG_PANEL, foreground=CLR_TEXT,
                        fieldbackground=BG_PANEL, borderwidth=0,
                        rowheight=28, font=("Consolas", 10))
        style.configure("V.Treeview.Heading",
                        background=BG_HEADER, foreground=CLR_CYAN,
                        font=("Consolas", 10, "bold"), relief="flat")
        style.map("V.Treeview",
                  background=[("selected", TV_SELECT)],
                  foreground=[("selected", CLR_TEXT)])

        wrapper = tk.Frame(parent, bg=CLR_BORDER, bd=1)
        wrapper.pack(fill="both", expand=True)

        cols = ("dest", "target", "device", "status", "rtt")
        self.tree = ttk.Treeview(wrapper, columns=cols, show="headings",
                                  style="V.Treeview", selectmode="none")

        for col, hdr, w in [
            ("dest",   "VLAN",      90),
            ("target", "TARGET IP", 135),
            ("device", "DEVICE",    150),
            ("status", "STATUS",    72),
            ("rtt",    "RTT",       68),
        ]:
            self.tree.heading(col, text=hdr, anchor="w")
            self.tree.column(col, width=w, minwidth=w, anchor="w")

        self.tree.pack(fill="both", expand=True)
        self.tree.tag_configure("reach",   foreground=CLR_GREEN)
        self.tree.tag_configure("block",   foreground=CLR_RED)
        self.tree.tag_configure("unknown", foreground=CLR_YELLOW)
        self.tree.tag_configure("self",    foreground=SELF_TAG)

        self.row_ids = {}
        for name in self.vlan_names:
            iid = self.tree.insert(
                "", "end",
                values=(name, self.vlans[name]["target"],
                        self.vlans[name]["label"], "???", "—"),
                tags=("unknown",),
            )
            self.row_ids[name] = iid

    def _build_matrix(self, parent):
        n  = len(self.vlan_names)
        px = self.CELL_PX
        lw = self.LABEL_W
        lh = self.LABEL_H

        wrapper = tk.Frame(parent, bg=CLR_BORDER, bd=1)
        wrapper.pack()

        self.canvas = tk.Canvas(
            wrapper, bg=BG_DARK,
            width=lw + n * px + 4,
            height=lh + n * px + 4,
            highlightthickness=0,
        )
        self.canvas.pack()

        for j, name in enumerate(self.vlan_names):
            self.canvas.create_text(
                lw + j * px + px // 2, lh // 2,
                text=name[:6], font=("Consolas", 7, "bold"),
                fill=CLR_MUTED, anchor="center",
            )

        for i, name in enumerate(self.vlan_names):
            self.canvas.create_text(
                lw - 6, lh + i * px + px // 2,
                text=name, font=("Consolas", 8, "bold"),
                fill=CLR_CYAN, anchor="e",
            )

        self.cell_rects = {}
        self.cell_texts = {}
        for i, frm in enumerate(self.vlan_names):
            for j, to in enumerate(self.vlan_names):
                x0 = lw + j * px + 2
                y0 = lh + i * px + 2
                x1 = x0 + px - 4
                y1 = y0 + px - 4
                r  = self.canvas.create_rectangle(
                    x0, y0, x1, y1, fill=CELL_GREY, outline=BG_DARK, width=1,
                )
                t  = self.canvas.create_text(
                    (x0 + x1) // 2, (y0 + y1) // 2,
                    text="?", font=("Consolas", 8), fill=CLR_MUTED,
                )
                self.cell_rects[(frm, to)] = r
                self.cell_texts[(frm, to)] = t

        self.matrix_tooltip = MatrixTooltip(self, self.canvas)

        legend = tk.Frame(parent, bg=BG_DARK)
        legend.pack(anchor="w", pady=(8, 0))
        for color, label in [(CLR_GREEN, "Reachable"), (CLR_RED, "Blocked"), (CLR_MUTED, "Untested")]:
            tk.Label(legend, text="■", font=("Consolas", 13),
                     fg=color, bg=BG_DARK).pack(side="left")
            tk.Label(legend, text=f" {label}    ", font=("Consolas", 9),
                     fg=CLR_MUTED, bg=BG_DARK).pack(side="left")

    def _build_footer(self, parent):
        bar = tk.Frame(parent, bg=BG_HEADER, height=32)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)

        self.stats_var = tk.StringVar(value="Coverage: 0/0   Reachable: 0   Blocked: 0")
        tk.Label(bar, textvariable=self.stats_var,
                 font=("Consolas", 10), fg=CLR_MUTED, bg=BG_HEADER,
                 ).pack(side="left", padx=20, pady=6)

    # ── Config Tab ────────────────────────────────────────────────────────────

    def _build_config_tab(self):
        for w in self.config_tab.winfo_children():
            w.destroy()

        outer = tk.Frame(self.config_tab, bg=BG_DARK)
        outer.pack(fill="both", expand=True, padx=24, pady=18)

        # ── VLAN list ──
        _section_label(outer, "VLAN DEFINITIONS")

        list_frame = tk.Frame(outer, bg=BG_DARK)
        list_frame.pack(fill="x")

        # Treeview
        tv_wrap = tk.Frame(list_frame, bg=CLR_BORDER, bd=1)
        tv_wrap.pack(side="left", fill="x", expand=True)

        style = ttk.Style()
        style.configure("Cfg.Treeview",
                        background=BG_PANEL, foreground=CLR_TEXT,
                        fieldbackground=BG_PANEL, borderwidth=0,
                        rowheight=30, font=("Consolas", 10))
        style.configure("Cfg.Treeview.Heading",
                        background=BG_HEADER, foreground=CLR_CYAN,
                        font=("Consolas", 10, "bold"), relief="flat")
        style.map("Cfg.Treeview",
                  background=[("selected", TV_SELECT)],
                  foreground=[("selected", CLR_TEXT)])

        self.cfg_tree = ttk.Treeview(
            tv_wrap,
            columns=("name", "subnet", "target", "label"),
            show="headings",
            style="Cfg.Treeview",
            selectmode="browse",
            height=8,
        )
        for col, hdr, w in [
            ("name",   "VLAN NAME",   120),
            ("subnet", "SUBNET",      140),
            ("target", "TARGET IP",   140),
            ("label",  "DEVICE LABEL",200),
        ]:
            self.cfg_tree.heading(col, text=hdr, anchor="w")
            self.cfg_tree.column(col, width=w, minwidth=w, anchor="w")

        vlan_sb = ttk.Scrollbar(tv_wrap, orient="vertical", command=self.cfg_tree.yview)
        self.cfg_tree.configure(yscrollcommand=vlan_sb.set)
        vlan_sb.pack(side="right", fill="y")
        self.cfg_tree.pack(fill="both", expand=True)
        self.cfg_tree.bind("<Double-1>", lambda _e: self._cfg_edit())

        self._cfg_refresh_tree()

        # Buttons column
        btn_col = tk.Frame(list_frame, bg=BG_DARK)
        btn_col.pack(side="left", fill="y", padx=(10, 0))

        for text, bg, cmd in [
            ("＋  Add",    CLR_GREEN,  self._cfg_add),
            ("✏  Edit",   CLR_CYAN,   self._cfg_edit),
            ("🗑  Delete", CLR_RED,    self._cfg_delete),
            ("↑  Up",     CLR_MUTED,  self._cfg_move_up),
            ("↓  Down",   CLR_MUTED,  self._cfg_move_down),
        ]:
            tk.Button(
                btn_col, text=text, bg=bg, fg=BG_DARK if bg != CLR_MUTED else CLR_TEXT,
                command=cmd,
                font=("Consolas", 10, "bold"),
                relief="flat", padx=14, pady=7,
                cursor="hand2", bd=0, width=12,
            ).pack(pady=(0, 8), anchor="n")

        # ── NIC | Ping | Apply  (1/2 | 1/4 | 1/4) ──
        tk.Frame(outer, bg=CLR_BORDER, height=1).pack(fill="x", pady=(18, 14))

        row2 = tk.Frame(outer, bg=BG_DARK)
        row2.pack(fill="x")
        row2.columnconfigure(0, weight=2)   # NIC  — 1/2
        row2.columnconfigure(1, weight=1)   # Ping — 1/2

        # ── NIC (col 0, 1/2) ──
        nic_col = tk.Frame(row2, bg=BG_DARK)
        nic_col.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        _section_label(nic_col, "NETWORK INTERFACE")

        nic_tv_wrap = tk.Frame(nic_col, bg=CLR_BORDER, bd=1)
        nic_tv_wrap.pack(fill="x")

        self.nic_tree = ttk.Treeview(
            nic_tv_wrap,
            columns=("name", "ip"),
            show="headings",
            style="Cfg.Treeview",
            selectmode="browse",
            height=5,
        )
        for col, hdr, w in [
            ("name", "ADAPTER",   380),
            ("ip",   "IP ADDRESS",150),
        ]:
            self.nic_tree.heading(col, text=hdr, anchor="w")
            self.nic_tree.column(col, width=w, minwidth=w, anchor="w")

        self.nic_tree.tag_configure("selected_nic", foreground=CLR_GREEN)
        nic_sb = ttk.Scrollbar(nic_tv_wrap, orient="vertical", command=self.nic_tree.yview)
        self.nic_tree.configure(yscrollcommand=nic_sb.set)
        nic_sb.pack(side="right", fill="y")
        self.nic_tree.pack(fill="x")
        self.nic_tree.bind("<ButtonRelease-1>", self._nic_clicked)
        self._nic_populate()
        _btn(nic_col, "↺  Refresh", BG_PANEL, CLR_CYAN,
             self._nic_refresh).pack(anchor="e", pady=(4, 0))

        # ── Ping Settings (col 1, 1/4) ──
        ping_col = tk.Frame(row2, bg=BG_DARK)
        ping_col.grid(row=0, column=1, sticky="nsew", padx=(0, 12))

        _section_label(ping_col, "PING SETTINGS")

        self._interval_var = tk.StringVar(value=str(self.config.get("ping_interval", 5)))
        self._timeout_var  = tk.StringVar(value=str(self.config.get("ping_timeout", 2)))
        self._count_var    = tk.StringVar(value=str(self.config.get("ping_count", 1)))

        ping_fields = tk.Frame(ping_col, bg=BG_DARK)
        ping_fields.pack(anchor="w", fill="x")

        for row, (label, var, unit) in enumerate([
            ("Sweep interval", self._interval_var, "s"),
            ("Ping timeout",   self._timeout_var,  "s"),
            ("Pings per host", self._count_var,     ""),
        ]):
            tk.Label(ping_fields, text=label, font=("Consolas", 10),
                     fg=CLR_MUTED, bg=BG_DARK, anchor="w",
                     ).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=6)
            tk.Entry(ping_fields, textvariable=var, font=("Consolas", 11, "bold"),
                     bg=BG_PANEL, fg=CLR_CYAN, insertbackground=CLR_TEXT,
                     relief="flat", bd=6, width=5, justify="center",
                     ).grid(row=row, column=1, pady=6)
            tk.Label(ping_fields, text=unit, font=("Consolas", 10),
                     fg=CLR_MUTED, bg=BG_DARK,
                     ).grid(row=row, column=2, sticky="w", padx=(4, 0), pady=6)

        # ── Appearance (under ping settings) ──
        tk.Frame(ping_col, bg=BG_DARK, height=14).pack(fill="x")
        _section_label(ping_col, "APPEARANCE")

        theme_choices = [
            ("system", "System (follow Windows at launch)"),
            ("light",  "Light"),
            ("dark",   "Dark"),
        ]
        value_to_label = {v: l for v, l in theme_choices}
        label_to_value = {l: v for v, l in theme_choices}

        self._theme_var = tk.StringVar(value=self.config.get("theme", "system"))
        theme_display = tk.StringVar(
            value=value_to_label.get(self._theme_var.get(), theme_choices[0][1])
        )

        # Style the combobox to match the dark-panel look. Both the collapsed
        # field and the dropdown listbox need separate treatment — the listbox
        # is a Tk widget under the hood and only honours option_add.
        style = ttk.Style()
        style.configure("Theme.TCombobox",
                        fieldbackground=BG_PANEL, background=BG_PANEL,
                        foreground=CLR_TEXT, arrowcolor=CLR_CYAN,
                        bordercolor=CLR_BORDER, lightcolor=CLR_BORDER,
                        darkcolor=CLR_BORDER, selectbackground=BG_PANEL,
                        selectforeground=CLR_TEXT)
        style.map("Theme.TCombobox",
                  fieldbackground=[("readonly", BG_PANEL)],
                  foreground=[("readonly", CLR_TEXT)])
        self.root.option_add("*TCombobox*Listbox.background", BG_PANEL)
        self.root.option_add("*TCombobox*Listbox.foreground", CLR_TEXT)
        self.root.option_add("*TCombobox*Listbox.selectBackground", CLR_CYAN)
        self.root.option_add("*TCombobox*Listbox.selectForeground", BG_DARK)
        self.root.option_add("*TCombobox*Listbox.font", ("Consolas", 10))

        theme_combo = ttk.Combobox(
            ping_col, textvariable=theme_display, state="readonly",
            values=[l for _, l in theme_choices],
            style="Theme.TCombobox", font=("Consolas", 10),
        )
        theme_combo.pack(anchor="w", fill="x", pady=(4, 6))
        theme_combo.bind(
            "<<ComboboxSelected>>",
            lambda _e: self._theme_var.set(label_to_value.get(theme_display.get(), "system")),
        )


    # ── Config actions ────────────────────────────────────────────────────────

    def _cfg_refresh_tree(self):
        self.cfg_tree.delete(*self.cfg_tree.get_children())
        for v in self.config.get("vlans", []):
            self.cfg_tree.insert(
                "", "end",
                values=(v["name"], format_subnet(v["subnet"]), v["target"], v.get("label", "")),
            )

    def _cfg_selected_index(self):
        sel = self.cfg_tree.selection()
        if not sel:
            return None
        children = self.cfg_tree.get_children()
        return children.index(sel[0])

    def _cfg_add(self):
        dlg = VlanDialog(self.root, "Add VLAN")
        self.root.wait_window(dlg)
        if dlg.result:
            vlans = self.config.setdefault("vlans", [])
            # Prevent duplicate names
            if any(v["name"] == dlg.result["name"] for v in vlans):
                messagebox.showwarning("Duplicate",
                    f'A VLAN named "{dlg.result["name"]}" already exists.', parent=self.root)
                return
            vlans.append(dlg.result)
            self._cfg_refresh_tree()

    def _cfg_edit(self):
        idx = self._cfg_selected_index()
        if idx is None:
            return
        current = self.config["vlans"][idx]
        dlg = VlanDialog(self.root, "Edit VLAN", initial=current)
        self.root.wait_window(dlg)
        if dlg.result:
            self.config["vlans"][idx] = dlg.result
            self._cfg_refresh_tree()

    def _cfg_delete(self):
        idx = self._cfg_selected_index()
        if idx is None:
            return
        name = self.config["vlans"][idx]["name"]
        if messagebox.askyesno("Delete VLAN",
                               f'Remove "{name}" from the list?', parent=self.root):
            del self.config["vlans"][idx]
            self._cfg_refresh_tree()

    def _cfg_move_up(self):
        idx = self._cfg_selected_index()
        if idx is None or idx == 0:
            return
        vlans = self.config["vlans"]
        vlans[idx - 1], vlans[idx] = vlans[idx], vlans[idx - 1]
        self._cfg_refresh_tree()
        children = self.cfg_tree.get_children()
        self.cfg_tree.selection_set(children[idx - 1])

    def _cfg_move_down(self):
        idx = self._cfg_selected_index()
        vlans = self.config.get("vlans", [])
        if idx is None or idx >= len(vlans) - 1:
            return
        vlans[idx], vlans[idx + 1] = vlans[idx + 1], vlans[idx]
        self._cfg_refresh_tree()
        children = self.cfg_tree.get_children()
        self.cfg_tree.selection_set(children[idx + 1])

    def _nic_populate(self):
        self.nic_tree.delete(*self.nic_tree.get_children())
        selected = self.config.get("selected_nic")

        for iface in get_network_interfaces():
            is_sel = selected and (iface["alias"] == selected or iface["ip"] == selected)
            tag = ("selected_nic",) if is_sel else ()
            self.nic_tree.insert("", "end",
                                 values=(iface["name"], iface["ip"]),
                                 tags=tag)

    def _nic_refresh(self):
        self._nic_populate()

    def _nic_clicked(self, _event):
        sel = self.nic_tree.selection()
        if not sel:
            return
        vals = self.nic_tree.item(sel[0], "values")
        # Store alias ("Wi-Fi") — stable across DHCP renewals
        alias = next(
            (i["alias"] for i in get_network_interfaces() if i["name"] == vals[0]),
            vals[1],   # fallback to IP if name not found
        )
        self.config["selected_nic"] = alias
        self._nic_populate()

    def _cfg_export_config(self):
        default_name = f"vlan_config_{datetime.now().strftime('%Y-%m-%d')}.json"
        filename = filedialog.asksaveasfilename(
            parent=self.root,
            title="Export VLAN Configuration",
            defaultextension=".json",
            initialfile=default_name,
            filetypes=[("JSON Config", "*.json"), ("All Files", "*.*")],
        )
        if not filename:
            return
        try:
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2)
            self.apply_msg.set(f"✓  Exported: {os.path.basename(filename)}")
            self.root.after(4000, lambda: self.apply_msg.set(""))
        except Exception as e:
            messagebox.showerror("Export failed",
                                 f"Could not export config:\n{e}", parent=self.root)

    def _cfg_import_config(self):
        filename = filedialog.askopenfilename(
            parent=self.root,
            title="Import VLAN Configuration",
            defaultextension=".json",
            filetypes=[("JSON Config", "*.json"), ("All Files", "*.*")],
        )
        if not filename:
            return
        try:
            with open(filename, "r", encoding="utf-8") as f:
                imported = json.load(f)
        except Exception as e:
            messagebox.showerror("Import failed",
                f"Could not read file:\n{e}", parent=self.root)
            return

        if not isinstance(imported, dict) or "vlans" not in imported:
            messagebox.showerror("Invalid config",
                "File does not contain a valid VLAN configuration.",
                parent=self.root)
            return

        existing = len(self.config.get("vlans", []))
        if existing and not messagebox.askyesno(
            "Replace configuration",
            f"This will replace your current {existing} VLAN(s) with "
            f"{len(imported.get('vlans', []))} imported VLAN(s).\n\nContinue?",
            parent=self.root):
            return

        for k, v in DEFAULT_CONFIG.items():
            imported.setdefault(k, v)
        self.config = imported
        save_config(self.config)

        # Refresh Config tab widgets from new values
        self._cfg_refresh_tree()
        self._nic_populate()
        self._interval_var.set(str(self.config.get("ping_interval", 5)))
        self._timeout_var.set(str(self.config.get("ping_timeout", 2)))
        self._count_var.set(str(self.config.get("ping_count", 1)))

        # Activate the imported config straight away
        self._cfg_apply()

    def _cfg_apply(self):
        try:
            self.config["ping_interval"] = max(1, int(self._interval_var.get()))
            self.config["ping_timeout"]  = max(1, int(self._timeout_var.get()))
            self.config["ping_count"]    = max(1, int(self._count_var.get()))
        except ValueError:
            messagebox.showwarning("Invalid value",
                                   "Ping settings must be whole numbers.", parent=self.root)
            return

        # Detect theme change — tkinter colors are baked in at widget construction,
        # so a palette swap requires restarting the process.
        old_theme = self.config.get("theme", "system")
        new_theme = self._theme_var.get() if hasattr(self, "_theme_var") else old_theme
        theme_changed = new_theme != old_theme
        self.config["theme"] = new_theme

        save_config(self.config)

        if theme_changed:
            if messagebox.askyesno(
                "Restart to apply theme",
                "The theme change will take effect after a restart.\n\n"
                "Restart the app now?",
                parent=self.root,
            ):
                self._restart_app()
                return

        # Restart sweep with new config
        self._derive_vlans()
        with self._lock:
            self.sweep_count  = 0
            self.paused       = False
            self._restart_requested = True

        self._build_monitor_tab()
        self.apply_msg.set(f"✓  Saved — sweep restarted with {len(self.vlan_names)} VLANs.")
        self.root.after(3000, lambda: self.apply_msg.set(""))
        self._show_tab("monitor")

    def _restart_app(self):
        """Relaunch the app process so a new theme palette is applied."""
        self.running = False
        self._snapshot_geometry()
        try:
            save_config(self.config)
            save_results(self.results)
        except Exception:
            pass
        try:
            # PyInstaller onefile sets _MEIPASS2 / _PYI_* env vars that tell a
            # child exe to reuse the parent's _MEI<random> extraction dir. When
            # the parent exits it tears that dir down, so a child that inherits
            # these vars races the cleanup and fails with "Failed to import
            # encodings module". Strip them so the child bootstraps cleanly
            # into its own extraction dir.
            env = os.environ.copy()
            for key in list(env.keys()):
                if key.startswith("_PYI") or key == "_MEIPASS2":
                    env.pop(key, None)
            if getattr(sys, "frozen", False):
                subprocess.Popen(
                    [sys.executable] + sys.argv[1:],
                    creationflags=subprocess.CREATE_NO_WINDOW,
                    env=env,
                    close_fds=True,
                )
            else:
                subprocess.Popen([sys.executable] + sys.argv, env=env)
        except Exception as exc:
            messagebox.showerror(
                "Restart failed",
                f"Could not relaunch:\n{exc}\n\n"
                "Please close and reopen the app to apply the new theme.",
                parent=self.root,
            )
            return
        self.root.destroy()

    # ── Monitor Controls ──────────────────────────────────────────────────────

    def _toggle_pause(self):
        self.paused = not self.paused
        if self.paused:
            self.pause_btn.config(text="▶  RESUME", bg=CLR_YELLOW)
            self.status_lbl.config(text="● PAUSED", fg=CLR_YELLOW)
        else:
            self.pause_btn.config(text="⏸  RUNNING", bg=CLR_GREEN)
            self.status_lbl.config(text="● ACTIVE", fg=CLR_GREEN)

    def _renew_ip(self):
        if not self.paused:
            self._toggle_pause()
        self.stats_var.set("⟳  Releasing IP address…")
        threading.Thread(target=self._renew_ip_worker, daemon=True).start()

    def _renew_ip_worker(self):
        alias = self.config.get("selected_nic")   # e.g. "Wi-Fi", or None for auto

        # Snapshot current IP so we can detect when it changes
        old_ip = None
        if alias:
            for iface in get_network_interfaces():
                if iface["alias"] == alias:
                    old_ip = iface["ip"]
                    break

        def run(args):
            subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                           creationflags=subprocess.CREATE_NO_WINDOW)

        self.root.after(0, lambda: self.stats_var.set("⟳  Releasing IP address…"))
        if alias:
            run(["ipconfig", "/release", alias])
        else:
            run(["ipconfig", "/release"])

        self.root.after(0, lambda: self.stats_var.set("⟳  Renewing IP address…"))
        if alias:
            run(["ipconfig", "/renew", alias])
        else:
            run(["ipconfig", "/renew"])

        self.root.after(0, lambda: self.stats_var.set("⟳  Waiting for new IP…"))
        new_ip = _wait_for_new_ip(alias, old_ip) if alias else None

        self.root.after(0, lambda: self._renew_ip_done(alias, new_ip))

    def _renew_ip_done(self, alias, new_ip):
        if new_ip:
            # Update status bar immediately — don't wait for the worker's next loop
            self.my_ip = new_ip
            with self._lock:
                vlans = dict(self.vlans)
            vlan, _ = detect_vlan([new_ip], vlans)
            self.current_vlan = vlan
            msg = f"✓  IP renewed → {new_ip}   │   Press Resume when ready"
        else:
            msg = "⚠  No IP assigned yet — check connection, then press Resume"

        if hasattr(self, "nic_tree"):
            self._nic_populate()

        self.stats_var.set(msg)

    def _clear_matrix(self):
        with self._lock:
            self.results     = {}
            self.sweep_count = 0
        save_results(self.results)

    def _do_export(self):
        default_name = f"VLAN_Report_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.pdf"
        filename = filedialog.asksaveasfilename(
            parent=self.root,
            title="Save VLAN Report",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF Report", "*.pdf"), ("All Files", "*.*")],
        )
        if not filename:
            return
        try:
            export_report_pdf(
                self.vlan_names, self.vlans, self.results,
                self.current_vlan, self.my_ip, filename,
            )
            orig = self.stats_var.get()
            self.stats_var.set(f"✓  Report saved: {os.path.basename(filename)}")
            self.root.after(3000, lambda: self.stats_var.set(orig))
        except Exception as e:
            messagebox.showerror("Export failed",
                                 f"Could not save report:\n{e}", parent=self.root)

    # ── Worker Thread ─────────────────────────────────────────────────────────

    def _start_worker(self):
        self._restart_requested = False
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        while self.running:
            self._restart_requested = False   # clear at the top of every iteration

            interval = self.config.get("ping_interval", 5)
            timeout  = self.config.get("ping_timeout",  2)
            count    = self.config.get("ping_count",    1)

            # Resolve selected NIC alias to its current live IP every sweep.
            # This picks up DHCP changes automatically without touching the saved config.
            selected_alias = self.config.get("selected_nic")
            ifaces    = get_network_interfaces()
            source_ip = next(
                (i["ip"] for i in ifaces if i["alias"] == selected_alias), None
            ) or next(
                (i["ip"] for i in ifaces if i["ip"] == selected_alias), None
            )
            local_ips = [source_ip] if source_ip else []

            with self._lock:
                vlans      = dict(self.vlans)
                vlan_names = list(self.vlan_names)

            self.current_vlan, self.my_ip = detect_vlan(local_ips, vlans)

            if self.paused or not vlan_names:
                time.sleep(1)
                continue

            self.sweep_count += 1

            for vlan_name, vlan in vlans.items():
                if not self.running:
                    return
                if self._restart_requested:
                    break   # restart the outer while-loop with new config
                key          = f"{self.current_vlan or 'UNKNOWN'}->{vlan_name}"
                reached, rtt = ping(vlan["target"], count=count, timeout=timeout,
                                    source_ip=source_ip)
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with self._lock:
                    entry = self.results.get(key) or {}
                    history = entry.get("history")
                    if not isinstance(history, list):
                        # Legacy-format entry — seed history from the single
                        # stored point so we don't lose prior context.
                        history = []
                        if "last" in entry:
                            history.append({
                                "ok":   entry.get("last"),
                                "rtt":  entry.get("rtt"),
                                "time": entry.get("time"),
                            })
                    history.append({"ok": reached, "rtt": rtt, "time": now_str})
                    history = history[-20:]   # rolling window
                    entry.update({
                        "last":    reached,
                        "rtt":     rtt,
                        "time":    now_str,
                        "from_ip": self.my_ip,
                        "history": history,
                    })
                    self.results[key] = entry
            else:
                # Sweep completed without a restart — do countdown
                save_results(self.results)

                for remaining in range(interval, 0, -1):
                    if not self.running:
                        return
                    if self._restart_requested:
                        break
                    while self.paused and self.running:
                        time.sleep(0.5)
                    self.countdown = remaining
                    time.sleep(1)

                self.countdown = 0

    # ── UI Refresh (main thread) ───────────────────────────────────────────────

    def _tick(self):
        if not self.running:
            return

        self.ts_var.set(datetime.now().strftime("%Y-%m-%d  %H:%M:%S"))

        if hasattr(self, "vlan_var"):
            self.vlan_var.set(self.current_vlan or "UNKNOWN")
            self.ip_var.set(self.my_ip)
            selected_nic = self.config.get("selected_nic")
            self.nic_var.set(selected_nic if selected_nic else "none")
            self.sweep_var.set(str(self.sweep_count))

            if self.paused:
                self.countdown_var.set("")
            elif self.countdown > 0:
                self.countdown_var.set(f"  Next sweep in {self.countdown}s")
            else:
                self.countdown_var.set("  Sweeping…" if self.vlan_names else "")

        if self.vlan_names:
            self._refresh_sweep_table()
            self._refresh_matrix()
            self._refresh_stats()

        self.root.after(500, self._tick)

    def _refresh_sweep_table(self):
        if not hasattr(self, "row_ids"):
            return
        source = self.current_vlan or "UNKNOWN"
        with self._lock:
            results = dict(self.results)

        for name in self.vlan_names:
            if name not in self.row_ids:
                continue
            entry  = results.get(f"{source}->{name}", {})
            state  = entry.get("last")
            rtt    = entry.get("rtt")
            rtt_s  = f"{rtt:.1f}ms" if rtt is not None else "—"
            is_you = (name == self.current_vlan)

            if is_you:
                status_s, tag = "THIS SUBNET", "self"
            elif state is True:
                status_s, tag = "REACH", "reach"
            elif state is False:
                status_s, tag = "BLOCK", "block"
            else:
                status_s, tag = "???", "unknown"

            self.tree.item(
                self.row_ids[name],
                values=(name, self.vlans[name]["target"],
                        self.vlans[name]["label"], status_s, rtt_s),
                tags=(tag,),
            )

    def _refresh_matrix(self):
        if not hasattr(self, "cell_rects"):
            return
        with self._lock:
            results = dict(self.results)

        for frm in self.vlan_names:
            for to in self.vlan_names:
                entry = results.get(f"{frm}->{to}")
                rect  = self.cell_rects.get((frm, to))
                txt   = self.cell_texts.get((frm, to))
                if rect is None:
                    continue

                if entry is None:
                    fill, sym, tf = CELL_GREY, "?", CLR_MUTED
                elif entry.get("last"):
                    rtt  = entry.get("rtt")
                    sym  = f"{rtt:.0f}" if rtt is not None else "✓"
                    fill, tf = CELL_GREEN, CLR_GREEN
                else:
                    fill, sym, tf = CELL_RED, "✗", CLR_RED

                self.canvas.itemconfig(rect, fill=fill)
                self.canvas.itemconfig(txt, text=sym, fill=tf)

    def _refresh_stats(self):
        if not hasattr(self, "stats_var"):
            return
        n       = len(self.vlan_names)
        total   = n * n
        with self._lock:
            results = dict(self.results)
        tested  = len(results)
        reach   = sum(1 for v in results.values() if v.get("last") is True)
        blocked = sum(1 for v in results.values() if v.get("last") is False)
        self.stats_var.set(
            f"Coverage: {tested}/{total}   Reachable: {reach}   Blocked: {blocked}"
        )

    # ── Shutdown ──────────────────────────────────────────────────────────────

    def on_close(self):
        self.running = False
        self._snapshot_geometry()
        save_config(self.config)
        save_results(self.results)
        self.root.destroy()

# ─── Entry Point ──────────────────────────────────────────────────────────────

def main():
    apply_palette(resolve_theme(load_config().get("theme", "system")))
    root = tk.Tk()
    app  = VlanTesterApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    main()
