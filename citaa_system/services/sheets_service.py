# -*- coding: utf-8 -*-
"""Spreadsheet operations service"""
from typing import List, Dict, Any, Optional
from datetime import datetime
import gspread
from gspread import Spreadsheet, Worksheet
from gspread.exceptions import WorksheetNotFound, SpreadsheetNotFound
import config
from auth.google_auth import get_auth_manager
from services.error_logger import get_error_logger


class SheetsService:
    def __init__(self):
        self._master_ss: Optional[Spreadsheet] = None

    @property
    def client(self) -> gspread.Client:
        return get_auth_manager().gspread_client

    def get_master_ss(self) -> Spreadsheet:
        if self._master_ss is None:
            try:
                ss_id = config.user_settings.master_ss_id
                self._master_ss = self.client.open_by_key(ss_id)
            except SpreadsheetNotFound:
                raise RuntimeError(f"Master SS not found: {ss_id}")
        return self._master_ss

    def reset_master_ss(self):
        """Reset cached master spreadsheet (call after changing SS ID)"""
        self._master_ss = None

    def get_or_create_sheet(self, name: str, headers: Optional[List[str]] = None) -> Worksheet:
        ss = self.get_master_ss()
        try:
            sheet = ss.worksheet(name)
        except WorksheetNotFound:
            sheet = ss.add_worksheet(title=name, rows=1000, cols=26)
            if headers:
                sheet.append_row(headers)
        return sheet

    def get_all_records(self, sheet_name: str) -> List[Dict[str, Any]]:
        try:
            sheet = self.get_master_ss().worksheet(sheet_name)
            return sheet.get_all_records()
        except Exception as e:
            get_error_logger().log_error("SheetsService", f"get_all_records({sheet_name})", e)
            return []

    def append_row(self, sheet_name: str, row_data: List[Any]) -> None:
        sheet = self.get_or_create_sheet(sheet_name)
        sheet.append_row(row_data, value_input_option="USER_ENTERED")

    def get_all_values(self, sheet_name: str) -> List[List[Any]]:
        try:
            sheet = self.get_master_ss().worksheet(sheet_name)
            return sheet.get_all_values()
        except Exception as e:
            get_error_logger().log_error("SheetsService", f"get_all_values({sheet_name})", e)
            return []

    def clear_and_update(self, sheet_name: str, data: List[List[Any]]) -> None:
        sheet = self.get_or_create_sheet(sheet_name)
        sheet.clear()
        if data:
            sheet.update('A1', data)

    def delete_row(self, sheet_name: str, row_index: int) -> None:
        sheet = self.get_master_ss().worksheet(sheet_name)
        sheet.delete_rows(row_index + 1)
    
    def update_row(self, sheet_name: str, row_index: int, row_data: List[Any]) -> None:
        sheet = self.get_master_ss().worksheet(sheet_name)
        # row_index is 0-based, but sheet rows are 1-based + header
        sheet.update(f'A{row_index + 2}', [row_data])

    # ============================================================
    # Clubs
    # ============================================================
    def get_clubs(self) -> List[Dict[str, str]]:
        return self.get_all_records(config.SHEET_CLUBS)

    def save_clubs(self, clubs: List[Dict[str, str]]) -> None:
        headers = ["ClubName", "Category", "Color", "DisplayName"]
        data = [headers] + [[c.get("ClubName", ""), c.get("Category", ""), 
                             c.get("Color", ""), c.get("DisplayName", "")] for c in clubs]
        self.clear_and_update(config.SHEET_CLUBS, data)
    
    def add_club(self, club: Dict[str, str]) -> None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.append_row(config.SHEET_CLUBS, [
            club.get("ClubName", ""), club.get("Category", ""),
            club.get("Color", ""), club.get("DisplayName", "")
        ])
    
    def delete_club(self, index: int) -> None:
        self.delete_row(config.SHEET_CLUBS, index)

    # ============================================================
    # Categories
    # ============================================================
    def get_categories(self) -> List[Dict[str, str]]:
        return self.get_all_records(config.SHEET_CATEGORIES)
    
    def save_categories(self, categories: List[Dict[str, str]]) -> None:
        headers = ["CategoryName", "Order"]
        data = [headers] + [[c.get("CategoryName", ""), c.get("Order", "")] for c in categories]
        self.clear_and_update(config.SHEET_CATEGORIES, data)

    # ============================================================
    # Members
    # ============================================================
    def get_members(self) -> List[Dict[str, str]]:
        return self.get_all_records(config.SHEET_MEMBERS)

    def save_members(self, members: List[Dict[str, str]]) -> None:
        headers = ["StudentID", "Name", "Department", "Role"]
        data = [headers] + [[m.get("StudentID", ""), m.get("Name", ""), 
                             m.get("Department", ""), m.get("Role", "")] for m in members]
        self.clear_and_update(config.SHEET_MEMBERS, data)
    
    def upsert_member(self, member: Dict[str, str]) -> None:
        """Add or update member by name"""
        members = self.get_members()
        name = member.get("Name", "")
        found = False
        for i, m in enumerate(members):
            if m.get("Name") == name:
                members[i] = member
                found = True
                break
        if not found:
            members.append(member)
        self.save_members(members)

    # ============================================================
    # Facilities
    # ============================================================
    def get_facilities(self) -> List[Dict[str, str]]:
        return self.get_all_records(config.SHEET_FACILITIES)

    def save_facilities(self, facilities: List[Dict[str, str]]) -> None:
        headers = ["FacilityID", "FacilityName"]
        data = [headers] + [[f.get("FacilityID", ""), f.get("FacilityName", "")] for f in facilities]
        self.clear_and_update(config.SHEET_FACILITIES, data)

    # ============================================================
    # Secretary Log
    # ============================================================
    def get_secretary_logs(self, facility: Optional[str] = None, date: Optional[str] = None) -> List[Dict]:
        records = self.get_all_records(config.SHEET_SECRETARY_LOG)
        if facility:
            records = [r for r in records if r.get("Facility") == facility]
        if date:
            records = [r for r in records if r.get("Date") == date]
        return records

    def add_secretary_log(self, log: Dict[str, Any]) -> None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.append_row(config.SHEET_SECRETARY_LOG, [
            log.get("Date", ""), log.get("Facility", ""), log.get("ClubName", ""),
            log.get("StartTime", ""), log.get("EndTime", ""), log.get("Note", ""), now
        ])
    
    def delete_secretary_log(self, index: int) -> None:
        self.delete_row(config.SHEET_SECRETARY_LOG, index)

    # ============================================================
    # Finance
    # ============================================================
    def get_finance_entries(self) -> List[Dict[str, Any]]:
        return self.get_all_records(config.SHEET_FINANCE)

    def add_finance_entry(self, entry: Dict[str, Any]) -> None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.append_row(config.SHEET_FINANCE, [
            entry.get("Date", ""), entry.get("Subject", ""), entry.get("Description", ""),
            entry.get("PaymentMethod", ""), entry.get("Income", 0), entry.get("Expense", 0),
            entry.get("ReimbursementStatus", "/"), now
        ])

    def update_finance_entry(self, index: int, entry: Dict[str, Any]) -> None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.update_row(config.SHEET_FINANCE, index, [
            entry.get("Date", ""), entry.get("Subject", ""), entry.get("Description", ""),
            entry.get("PaymentMethod", ""), entry.get("Income", 0), entry.get("Expense", 0),
            entry.get("ReimbursementStatus", "/"), now
        ])

    def delete_finance_entry(self, index: int) -> None:
        self.delete_row(config.SHEET_FINANCE, index + 1)

    # ============================================================
    # Attendance & Weekday Assignment
    # ============================================================
    def get_attendance(self) -> List[Dict[str, Any]]:
        return self.get_all_records(config.SHEET_ATTENDANCE)

    def add_attendance(self, record: Dict[str, Any]) -> None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.append_row(config.SHEET_ATTENDANCE, [
            record.get("Date", ""), record.get("MemberName", ""),
            record.get("Status", ""), record.get("Period", ""), now
        ])

    def get_weekday_assignments(self) -> Dict[str, Dict[str, List[str]]]:
        """Get weekday assignments organized by period and day"""
        records = self.get_all_records(config.SHEET_WEEKDAY_ASSIGN)
        result = {}
        for r in records:
            period = r.get("Period", "")
            day = r.get("Weekday", "")
            members = r.get("Members", "")
            if period not in result:
                result[period] = {}
            result[period][day] = [m.strip() for m in members.split(",") if m.strip()]
        return result

    def save_weekday_assignments(self, assignments: Dict[str, Dict[str, List[str]]]) -> None:
        headers = ["Period", "Weekday", "Members"]
        data = [headers]
        for period, days in assignments.items():
            for day, members in days.items():
                data.append([period, day, ", ".join(members)])
        self.clear_and_update(config.SHEET_WEEKDAY_ASSIGN, data)

    # ============================================================
    # External Logs
    # ============================================================
    def get_external_logs(self, club: Optional[str] = None, log_type: Optional[str] = None) -> List[Dict[str, Any]]:
        records = self.get_all_records(config.SHEET_EXTERNAL_LOG)
        if club:
            records = [r for r in records if r.get("ClubName") == club]
        if log_type:
            records = [r for r in records if r.get("LogType") == log_type]
        return records

    def add_external_log(self, log: Dict[str, Any]) -> None:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.append_row(config.SHEET_EXTERNAL_LOG, [
            log.get("No", ""), log.get("LogType", ""), log.get("HasScan", ""),
            log.get("HasActivity", ""), log.get("HasReport", ""), 
            log.get("Period", ""), log.get("HasMatch", ""),
            log.get("TournamentName", ""), log.get("HasOvernight", ""),
            log.get("Organizer", ""), log.get("Participants", ""),
            log.get("TeamResult", ""), log.get("IndividualResult", ""),
            log.get("ClubName", ""), now
        ])

    # ============================================================
    # Editorial: Required Items, Passwords, Advisors
    # ============================================================
    def get_required_items(self) -> List[Dict[str, Any]]:
        return self.get_all_records(config.SHEET_REQUIRED_ITEMS)
    
    def save_required_items(self, items: List[Dict[str, Any]]) -> None:
        headers = ["ClubName", "StudentPhone", "GuarantorPhone", "Address"]
        data = [headers] + [[
            i.get("ClubName", ""),
            i.get("StudentPhone", ""),
            i.get("GuarantorPhone", ""),
            i.get("Address", "")
        ] for i in items]
        self.clear_and_update(config.SHEET_REQUIRED_ITEMS, data)
    
    def get_passwords(self) -> List[Dict[str, str]]:
        return self.get_all_records(config.SHEET_PASSWORDS)
    
    def save_passwords(self, passwords: List[Dict[str, str]]) -> None:
        headers = ["ClubName", "Password"]
        data = [headers] + [[p.get("ClubName", ""), p.get("Password", "")] for p in passwords]
        self.clear_and_update(config.SHEET_PASSWORDS, data)
    
    def get_advisors(self) -> List[Dict[str, str]]:
        return self.get_all_records(config.SHEET_ADVISORS)
    
    def save_advisors(self, advisors: List[Dict[str, str]]) -> None:
        headers = ["ClubName", "Director", "Advisor", "Coach", "CoachSub"]
        data = [headers] + [[
            a.get("ClubName", ""),
            a.get("Director", ""),
            a.get("Advisor", ""),
            a.get("Coach", ""),
            a.get("CoachSub", "")
        ] for a in advisors]
        self.clear_and_update(config.SHEET_ADVISORS, data)

    # ============================================================
    # Bookmarks (Editorial URLs)
    # ============================================================
    def get_bookmarks(self) -> List[Dict[str, str]]:
        return self.get_all_records(config.SHEET_BOOKMARKS)

    def save_bookmarks(self, bookmarks: List[Dict[str, str]]) -> None:
        headers = ["Name", "URL"]
        data = [headers] + [[b.get("Name", ""), b.get("URL", "")] for b in bookmarks]
        self.clear_and_update(config.SHEET_BOOKMARKS, data)

    # ============================================================
    # Study Weeks (Admin)
    # ============================================================
    def get_study_weeks(self) -> List[Dict[str, str]]:
        return self.get_all_records(config.SHEET_STUDY_WEEKS)
    
    def save_study_weeks(self, weeks: List[Dict[str, str]]) -> None:
        headers = ["Period", "StartDate", "EndDate"]
        data = [headers] + [[w.get("Period", ""), w.get("StartDate", ""), w.get("EndDate", "")] for w in weeks]
        self.clear_and_update(config.SHEET_STUDY_WEEKS, data)


_sheets_service: Optional[SheetsService] = None

def get_sheets_service() -> SheetsService:
    global _sheets_service
    if _sheets_service is None:
        _sheets_service = SheetsService()
    return _sheets_service
