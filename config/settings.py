# -*- coding: utf-8 -*-
"""User settings: PDF font path and Excel sheet name."""

import json
import os
from typing import Any, Dict, Optional

# Defaults (aligned with constants)
try:
    from constants import FONT_PATH as DEFAULT_FONT_PATH, SAYFA_ADI as DEFAULT_SHEET_NAME
except ImportError:
    DEFAULT_FONT_PATH = "C:/Windows/Fonts/arial.ttf"
    DEFAULT_SHEET_NAME = "Coaching Schedule"

CONFIG_DIR = "config"
CONFIG_FILE = "ayarlar.json"


def _config_path(base_dir: str) -> str:
    return os.path.join(base_dir, CONFIG_DIR, CONFIG_FILE)


def load_settings(base_dir: Optional[str] = None) -> Dict[str, Any]:
    """Load settings. Returns defaults when missing or invalid."""
    base_dir = base_dir or os.getcwd()
    path = _config_path(base_dir)
    if not os.path.exists(path):
        return {"pdf_font_path": DEFAULT_FONT_PATH, "excel_sheet_name": DEFAULT_SHEET_NAME, "theme": "Gray (default)"}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {"pdf_font_path": DEFAULT_FONT_PATH, "excel_sheet_name": DEFAULT_SHEET_NAME, "theme": "Gray (default)"}
    return {
        "pdf_font_path": data.get("pdf_font_path", DEFAULT_FONT_PATH),
        "excel_sheet_name": data.get("excel_sheet_name", DEFAULT_SHEET_NAME),
        "theme": data.get("theme", "gri"),
    }


def save_settings(base_dir: Optional[str], settings: Dict[str, Any]) -> None:
    """Save settings."""
    base_dir = base_dir or os.getcwd()
    dirpath = os.path.join(base_dir, CONFIG_DIR)
    os.makedirs(dirpath, exist_ok=True)
    path = _config_path(base_dir)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2, ensure_ascii=False)


def get_pdf_font_path(base_dir: Optional[str] = None) -> str:
    return load_settings(base_dir).get("pdf_font_path", DEFAULT_FONT_PATH)


def get_sheet_name(base_dir: Optional[str] = None) -> str:
    return load_settings(base_dir).get("excel_sheet_name", DEFAULT_SHEET_NAME)


def get_theme(base_dir: Optional[str] = None) -> str:
    return load_settings(base_dir).get("theme", "Gray (default)")
