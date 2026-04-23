# -*- coding: utf-8 -*-
"""Google Authentication Manager with Enhanced UI Support"""
from typing import Optional, Tuple
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build, Resource
import gspread
import config
import webbrowser
import threading
import http.server
import socketserver
import urllib.parse


class AuthCompletionHandler(http.server.SimpleHTTPRequestHandler):
    """Custom handler to show a nice completion page"""
    
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/html; charset=utf-8')
        self.end_headers()
        
        html = '''<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>認証完了 - CITAA System</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', 'Hiragino Sans', 'Yu Gothic', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .container {
            background: white;
            border-radius: 24px;
            padding: 60px 80px;
            box-shadow: 0 25px 80px rgba(0,0,0,0.3);
            text-align: center;
            max-width: 600px;
            animation: fadeIn 0.5s ease-out;
        }
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        .icon {
            width: 120px;
            height: 120px;
            background: linear-gradient(135deg, #22c55e 0%, #16a34a 100%);
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 30px;
            animation: pulse 2s infinite;
        }
        @keyframes pulse {
            0%, 100% { transform: scale(1); }
            50% { transform: scale(1.05); }
        }
        .icon svg {
            width: 60px;
            height: 60px;
            fill: white;
        }
        h1 {
            font-size: 32px;
            color: #1f2937;
            margin-bottom: 20px;
            font-weight: 700;
        }
        .message {
            font-size: 20px;
            color: #4b5563;
            line-height: 1.8;
            margin-bottom: 30px;
        }
        .sub-message {
            font-size: 16px;
            color: #6b7280;
            background: #f3f4f6;
            padding: 20px;
            border-radius: 12px;
            margin-top: 20px;
        }
        .app-name {
            font-size: 14px;
            color: #9ca3af;
            margin-top: 30px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="icon">
            <svg viewBox="0 0 24 24">
                <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
            </svg>
        </div>
        <h1>✨ ログインが完了しました ✨</h1>
        <p class="message">
            認証が正常に完了いたしました。<br>
            <strong>このページは閉じて問題ございません。</strong>
        </p>
        <div class="sub-message">
            📱 CITAA System アプリケーションに戻り、<br>
            引き続きご利用ください。
        </div>
        <p class="app-name">千葉工業大学体育会本部 統合管理システム</p>
    </div>
</body>
</html>'''
        self.wfile.write(html.encode('utf-8'))
    
    def log_message(self, format, *args):
        pass  # Suppress server logs


class GoogleAuthManager:
    def __init__(self):
        self._credentials: Optional[Credentials] = None
        self._sheets_service: Optional[Resource] = None
        self._drive_service: Optional[Resource] = None
        self._gspread_client: Optional[gspread.Client] = None

    @property
    def is_authenticated(self) -> bool:
        return self._credentials is not None and self._credentials.valid

    @property
    def credentials(self) -> Optional[Credentials]:
        return self._credentials

    @property
    def sheets_service(self) -> Resource:
        if self._sheets_service is None:
            if not self.is_authenticated:
                raise RuntimeError("Not authenticated")
            self._sheets_service = build("sheets", "v4", credentials=self._credentials)
        return self._sheets_service

    @property
    def drive_service(self) -> Resource:
        if self._drive_service is None:
            if not self.is_authenticated:
                raise RuntimeError("Not authenticated")
            self._drive_service = build("drive", "v3", credentials=self._credentials)
        return self._drive_service

    @property
    def gspread_client(self) -> gspread.Client:
        if self._gspread_client is None:
            if not self.is_authenticated:
                raise RuntimeError("Not authenticated")
            self._gspread_client = gspread.authorize(self._credentials)
        return self._gspread_client

    def authenticate(self, force_new: bool = False) -> Tuple[bool, str]:
        try:
            if not config.CREDENTIALS_FILE.exists():
                return False, "credentials.json not found"
            creds = None
            if not force_new and config.TOKEN_FILE.exists():
                try:
                    creds = Credentials.from_authorized_user_file(
                        str(config.TOKEN_FILE), config.SCOPES
                    )
                except Exception:
                    creds = None
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception:
                    creds = None
            if not creds or not creds.valid:
                flow = InstalledAppFlow.from_client_secrets_file(
                    str(config.CREDENTIALS_FILE), config.SCOPES
                )
                creds = flow.run_local_server(
                    port=0,
                    success_message="認証が完了しました。このウィンドウを閉じてください。",
                    open_browser=True
                )
                config.AUTH_DIR.mkdir(exist_ok=True)
                with open(config.TOKEN_FILE, "w", encoding="utf-8") as f:
                    f.write(creds.to_json())
            self._credentials = creds
            self._sheets_service = None
            self._drive_service = None
            self._gspread_client = None
            return True, "OK"
        except Exception as e:
            return False, str(e)

    def logout(self) -> Tuple[bool, str]:
        try:
            if config.TOKEN_FILE.exists():
                config.TOKEN_FILE.unlink()
            self._credentials = None
            self._sheets_service = None
            self._drive_service = None
            self._gspread_client = None
            return True, "Logged out"
        except Exception as e:
            return False, str(e)

    def get_user_info(self) -> Optional[dict]:
        if not self.is_authenticated:
            return None
        try:
            about = self.drive_service.about().get(fields="user").execute()
            return about.get("user", {})
        except Exception:
            return None


_auth_manager: Optional[GoogleAuthManager] = None

def get_auth_manager() -> GoogleAuthManager:
    global _auth_manager
    if _auth_manager is None:
        _auth_manager = GoogleAuthManager()
    return _auth_manager
