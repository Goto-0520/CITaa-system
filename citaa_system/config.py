# -*- coding: utf-8 -*-
"""CITAA System Configuration"""
import json
from pathlib import Path
from typing import Optional

BASE_DIR = Path(__file__).resolve().parent
ASSETS_DIR = BASE_DIR / "assets"
AUTH_DIR = BASE_DIR / "auth"
CACHE_DIR = BASE_DIR / ".cache"
LOG_DIR = BASE_DIR / "logs"
USER_SETTINGS_FILE = BASE_DIR / "user_settings.json"

CREDENTIALS_FILE = AUTH_DIR / "credentials.json"
TOKEN_FILE = AUTH_DIR / "token.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Default Master Spreadsheet ID
DEFAULT_MASTER_SS_ID = "13ps08QUsjIfmm8xvi-iRu1pDZMnR92ygCShAvxoMdi8"

APP_TITLE = "CITAA System"
APP_WIDTH = 1400
APP_HEIGHT = 900

# ============================================================
# Sheet Names
# ============================================================
SHEET_SETTINGS = "Settings"
SHEET_CLUBS = "Clubs"
SHEET_MEMBERS = "Members"
SHEET_FACILITIES = "Facilities"
SHEET_SECRETARY_LOG = "SecretaryLog"
SHEET_FINANCE = "Finance"
SHEET_ATTENDANCE = "Attendance"
SHEET_WEEKDAY_ASSIGN = "WeekdayAssign"
SHEET_EXTERNAL_LOG = "ExternalLog"
SHEET_BOOKMARKS = "Bookmarks"
SHEET_ERROR_LOG = "ErrorLog"
SHEET_REQUIRED_ITEMS = "RequiredItems"
SHEET_PASSWORDS = "Passwords"
SHEET_ADVISORS = "Advisors"
SHEET_CATEGORIES = "Categories"
SHEET_STUDY_WEEKS = "StudyWeeks"

# ============================================================
# Colors - Monochrome Base + Accent
# ============================================================
# Base monochrome colors
COLOR_WHITE = "#ffffff"
COLOR_BLACK = "#1f2937"
COLOR_GRAY_50 = "#f9fafb"
COLOR_GRAY_100 = "#f3f4f6"
COLOR_GRAY_200 = "#e5e7eb"
COLOR_GRAY_300 = "#d1d5db"
COLOR_GRAY_400 = "#9ca3af"
COLOR_GRAY_500 = "#6b7280"
COLOR_GRAY_600 = "#4b5563"
COLOR_GRAY_700 = "#374151"
COLOR_GRAY_800 = "#1f2937"
COLOR_GRAY_900 = "#111827"

# Default accent color (CITAA purple/indigo)
DEFAULT_ACCENT = "#6366f1"

# Department card border colors
DEPT_COLORS = {
    "secretary": "#1e3a5f",
    "finance": "#f59e0b",
    "general": "#6b7280",
    "external": "#f97316",
    "editorial": "#ec4899",
    "event": "#06b6d4",
}

# Club colors for timeline (predefined)
CLUB_COLORS = [
    "#6366f1", "#8b5cf6", "#ec4899", "#ef4444", "#f97316",
    "#f59e0b", "#84cc16", "#22c55e", "#14b8a6", "#06b6d4",
    "#3b82f6", "#a855f7", "#d946ef", "#f43f5e", "#fb923c",
    "#facc15", "#a3e635", "#4ade80", "#2dd4bf", "#22d3ee",
    "#60a5fa", "#c084fc", "#f472b6", "#fb7185", "#fdba74",
    "#fde047", "#bef264", "#86efac", "#5eead4", "#67e8f9",
    "#93c5fd", "#d8b4fe", "#f9a8d4", "#fda4af", "#fed7aa",
    "#fef08a", "#d9f99d",
]

# ============================================================
# Time Settings
# ============================================================
TIMELINE_START_HOUR = 6
TIMELINE_END_HOUR = 20
TIMELINE_INTERVAL_MINUTES = 30

# ============================================================
# Period Settings for General Department
# ============================================================
PERIODS = [
    ("zenki", "前期"),
    ("kouki", "後期"),
    ("summer1", "夏季休業1"),
    ("summer2", "夏季休業2"),
    ("summer3", "夏季休業3"),
    ("winter1", "冬季休業1"),
    ("winter2", "冬季休業2"),
    ("winter3", "冬季休業3"),
]

WEEKDAYS = [
    ("mon", "月"),
    ("tue", "火"),
    ("wed", "水"),
    ("thu", "木"),
    ("fri", "金"),
    ("sat", "土"),
    ("sun", "日"),
]

# ============================================================
# Default Facilities
# ============================================================
DEFAULT_FACILITIES = [
    "体育館・トレーニングルーム",
    "陸上競技場・野球場",
    "ビーチバレー、ハンドボールコート・テニスコート",
    "ラグビー場・サッカー場",
    "武道館（剣道場・柔道場）",
    "多目的ルーム",
    "武道場（空手道場）・部室棟会議室",
    "射撃場・弓道場",
    "屋内練習場",
    "新習志野野球場",
]

# ============================================================
# Attendance Status Options
# ============================================================
ATTENDANCE_STATUS = [
    ("present", "出席"),
    ("absent", "欠席"),
    ("late", "遅刻"),
    ("early_leave", "早退"),
    ("mourning", "忌引等"),
]

# ============================================================
# Reimbursement Status Options (Finance)
# ============================================================
REIMBURSEMENT_STATUS = [
    ("/", "該当なし"),
    ("unreimbursed", "未返金"),
    ("reimbursed", "返金済"),
]

# ============================================================
# Payment Methods
# ============================================================
PAYMENT_METHODS = [
    ("cash", "現金"),
    ("bank", "通帳"),
]

# ============================================================
# User Settings Management
# ============================================================
class UserSettings:
    def __init__(self):
        self._settings = self._load()
    
    def _load(self) -> dict:
        if USER_SETTINGS_FILE.exists():
            try:
                with open(USER_SETTINGS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                pass
        return {
            "accent_color": DEFAULT_ACCENT,
            "master_ss_id": DEFAULT_MASTER_SS_ID,
            "onedrive_path": "",
            "printer_name": "",
            "scanner_url": "http://localhost:3000",
            "gemini_api_key": "",
            "system_password": "",
        }
    
    def _save(self):
        with open(USER_SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(self._settings, f, ensure_ascii=False, indent=2)
    
    @property
    def accent_color(self) -> str:
        return self._settings.get("accent_color", DEFAULT_ACCENT)
    
    @accent_color.setter
    def accent_color(self, value: str):
        self._settings["accent_color"] = value
        self._save()
    
    @property
    def master_ss_id(self) -> str:
        return self._settings.get("master_ss_id", DEFAULT_MASTER_SS_ID)
    
    @master_ss_id.setter
    def master_ss_id(self, value: str):
        self._settings["master_ss_id"] = value
        self._save()
    
    @property
    def onedrive_path(self) -> str:
        return self._settings.get("onedrive_path", "")
    
    @onedrive_path.setter
    def onedrive_path(self, value: str):
        self._settings["onedrive_path"] = value
        self._save()
    
    @property
    def printer_name(self) -> str:
        return self._settings.get("printer_name", "")
    
    @printer_name.setter
    def printer_name(self, value: str):
        self._settings["printer_name"] = value
        self._save()
    
    @property
    def scanner_url(self) -> str:
        return self._settings.get("scanner_url", "http://localhost:3000")
    
    @scanner_url.setter
    def scanner_url(self, value: str):
        self._settings["scanner_url"] = value
        self._save()
    
    @property
    def gemini_api_key(self) -> str:
        return self._settings.get("gemini_api_key", "")
    
    @gemini_api_key.setter
    def gemini_api_key(self, value: str):
        self._settings["gemini_api_key"] = value
        self._save()
    
    @property
    def system_password(self) -> str:
        return self._settings.get("system_password", "")
    
    @system_password.setter
    def system_password(self, value: str):
        self._settings["system_password"] = value
        self._save()

user_settings = UserSettings()

def ensure_directories():
    for d in [CACHE_DIR, AUTH_DIR, LOG_DIR, ASSETS_DIR]:
        d.mkdir(parents=True, exist_ok=True)
