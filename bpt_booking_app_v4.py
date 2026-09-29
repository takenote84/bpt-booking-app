#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""BPT 부킹 선반입 관리 — GUI 앱 (v4: 작업 결과 보고서 + Outlook Draft 저장 추가)"""

import sys, os, threading, queue, time, datetime, re, logging, subprocess, json
from pathlib import Path
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

# ── 자동 패키지 설치 ──
def ensure(pkg, import_name=None):
    name = import_name or pkg
    try:
        __import__(name)
    except ImportError:
        # --break-system-packages는 Windows pip에서 오류 발생하므로 제거
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", pkg])

ensure("playwright")
ensure("pywinauto")
ensure("pyautogui")
ensure("pyperclip")
ensure("openpyxl")                       # v4: 엑셀 보고서 저장용
ensure("pywin32", "win32com.client")     # v4: Outlook 메일 작성/Draft 저장용

# playwright 브라우저 설치 (최초 1회)
def ensure_playwright_browsers():
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            # chromium 실행 가능 여부 확인
            try:
                b = pw.chromium.launch(headless=True)
                b.close()
                return  # 이미 설치됨
            except Exception:
                pass
        # 설치 필요
        print("Playwright 브라우저 설치 중... (최초 1회)")
        subprocess.check_call([sys.executable, "-m", "playwright", "install", "chromium"])
        print("Playwright 브라우저 설치 완료")
    except Exception as e:
        print(f"Playwright 브라우저 설치 확인 오류: {e}")

ensure_playwright_browsers()

from playwright.sync_api import sync_playwright
from pywinauto import Desktop
import pyautogui, pyperclip
from typing import Optional
from dataclasses import dataclass, field

# ── 설정 ──
BPT_URL     = "https://info.bptc.co.kr/content/index.jsp"
BOOKING_URL = "https://info.bptc.co.kr/content/yw/frame/pregate_control_frame_yw_kr.jsp?p_id=PRGA_YD_KR&snb_num=3&snb_div=service"
CARRIER     = "SKR"

# v4: 작업 결과 보고서 수신자 / 참조(향후 확장용 — 필요 시 이메일 주소를 넣어 사용)
REPORT_MAIL_TO = "bcsto@sinokor.co.kr"
REPORT_MAIL_CC = ""

def _get_app_dir():
    """실행 파일 위치 반환 — .py 직접 실행과 PyInstaller .exe 모두 대응"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def _load_config():
    """config.json에서 BPT 계정 로드. 없으면 빈 양식을 만든다 — 계정은 소스에 넣지 않는다."""
    config_path = os.path.join(_get_app_dir(), "config.json")
    default = {"BPT_ID": "", "BPT_PW": ""}
    if os.path.exists(config_path):
        try:
            with open(config_path, encoding="utf-8") as f:
                data = json.load(f)
            return data.get("BPT_ID", default["BPT_ID"]), data.get("BPT_PW", default["BPT_PW"])
        except Exception:
            pass
    else:
        # 최초 실행 시 config.json 생성
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(default, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    return default["BPT_ID"], default["BPT_PW"]

BPT_ID, BPT_PW = _load_config()

@dataclass
class Booking:
    booking_no: str
    req_qty: int
    recv_qty: Optional[int]
    confirm: str
    etd: Optional[datetime.date] = None
    action: str = ""
    new_qty: int = 0
    fail_reason: str = ""   # v4: 삭제/수정 실패 사유 (보고서용)

def get_tomorrow():
    return datetime.date.today() + datetime.timedelta(days=1)

# ── ETD 캐시 ──────────────────────────────────────

CACHE_FILE = os.path.join(_get_app_dir(), "etd_cache.json")

def load_etd_cache() -> dict:
    """오늘 날짜 캐시 로드. 날짜가 다르면 빈 dict 반환"""
    try:
        if os.path.exists(CACHE_FILE):
            import json
            with open(CACHE_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if data.get("date") == str(datetime.date.today()):
                return data.get("etd_map", {})
    except Exception:
        pass
    return {}

def save_etd_cache(etd_map: dict):
    """오늘 날짜와 함께 캐시 저장"""
    try:
        import json
        data = {"date": str(datetime.date.today()), "etd_map": etd_map}
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def parse_etd(s: str) -> Optional[datetime.date]:
    digits = ''.join(filter(str.isdigit, s.strip()))
    if len(digits) >= 8:
        try:
            return datetime.date(int(digits[:4]), int(digits[4:6]), int(digits[6:8]))
        except ValueError:
            pass
    return None


# ══════════════════════════════════════════════════════
# v4: 작업 결과 보고서 & Outlook 메일 Draft 저장
#     (기존 로직과 분리된 독립 함수들 — App 클래스에서 Hook 방식으로만 호출)
# ══════════════════════════════════════════════════════

@dataclass
class Report:
    all_bookings: list      # 섹션① 검색된 전체 부킹
    targets: list           # 섹션② ETD 지난 삭제/수정 대상 (=처리 시도한 대상)
    summary: dict = field(default_factory=dict)


def generate_report(all_bookings, targets) -> Report:
    """처리 완료 후 all_bookings(전체 조회 결과)와 targets(처리 대상)로
    보고서 데이터와 Summary를 만든다."""
    del_ok = sum(1 for b in targets if b.action == "done_delete")
    mod_ok = sum(1 for b in targets if b.action == "done_modify")
    fail   = sum(1 for b in targets if b.action in ("delete", "modify"))
    summary = {
        "total": len(all_bookings),
        "etd_over": len(targets),
        "del_ok": del_ok,
        "mod_ok": mod_ok,
        "fail": fail,
    }
    return Report(all_bookings=list(all_bookings), targets=list(targets), summary=summary)


def _action_label(bk: "Booking") -> str:
    """Booking.action 값을 보고서용 한글 라벨로 변환 (섹션①/② 공용)"""
    if bk.action == "modify":
        return f"수정 → {bk.new_qty} van"
    if bk.action == "delete":
        return "삭제"
    if bk.action == "same":
        return ""
    if bk.action == "skip":
        return "스킵 (ETD 여유)"
    if bk.action == "done_delete":
        return "삭제 완료"
    if bk.action == "done_modify":
        return f"수정 완료 → {bk.new_qty} van"
    return "ETD 미확인" if not bk.action else bk.action


def _result_label(bk: "Booking") -> str:
    """섹션③ 처리결과 전용 라벨 (성공/실패 구분)"""
    if bk.action == "done_delete":
        return "삭제 성공"
    if bk.action == "done_modify":
        return "수정 성공"
    if bk.action == "delete":
        return "삭제 실패"
    if bk.action == "modify":
        return "수정 실패"
    return bk.action or "-"


def export_excel_report(report: "Report") -> Optional[str]:
    """보고서를 xlsx로 저장. 성공 시 파일 경로, 실패 시 None 반환.
    (로그 출력은 호출부에서 반환값을 보고 처리 — 기존 self.log 흐름과 분리)"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment

        wb = Workbook()
        header_fill = PatternFill(start_color="0D3B6E", end_color="0D3B6E", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)

        def write_table(ws, headers, rows):
            ws.append(headers)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
            for row in rows:
                ws.append(row)
            for col in ws.columns:
                max_len = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = max(10, max_len + 2)

        # 섹션① 검색된 전체 부킹
        ws1 = wb.active
        ws1.title = "전체 부킹"
        write_table(
            ws1,
            ["Booking No", "Request Qty", "Receive Qty", "Confirm", "ETD", "Action"],
            [[b.booking_no, b.req_qty, b.recv_qty if b.recv_qty is not None else "-",
              b.confirm, str(b.etd) if b.etd else "-", _action_label(b)]
             for b in report.all_bookings],
        )

        # 섹션② ETD가 지난 삭제/수정 대상
        ws2 = wb.create_sheet("ETD 지난 대상")
        write_table(
            ws2,
            ["Booking No", "Request Qty", "Receive Qty", "Confirm", "ETD", "Action"],
            [[b.booking_no, b.req_qty, b.recv_qty if b.recv_qty is not None else "-",
              b.confirm, str(b.etd) if b.etd else "-", _action_label(b)]
             for b in report.targets],
        )

        # 섹션③ 실제 처리 완료 결과 (성공/실패 모두 포함)
        ws3 = wb.create_sheet("처리 결과")
        write_table(
            ws3,
            ["Booking No", "구분", "결과", "실패 사유"],
            [[b.booking_no,
              "삭제" if b.action in ("delete", "done_delete") else "수정",
              _result_label(b),
              b.fail_reason or "-"]
             for b in report.targets],
        )

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"BPT_Delete_Report_{ts}.xlsx"
        path = os.path.join(_get_app_dir(), filename)
        wb.save(path)
        return path
    except Exception as e:
        print(f"[보고서] 엑셀 저장 실패: {e}")
        return None


def _mail_td(text, align="left", extra=""):
    """Outlook 호환 td — Inline CSS + 회색 border, Calibri/맑은 고딕 10.5pt"""
    style = ("border:1px solid #cccccc;font-family:Calibri,'맑은 고딕',sans-serif;"
              f"font-size:10.5pt;color:#333333;padding:6px 10px;text-align:{align};{extra}")
    return f'<td style="{style}">{text}</td>'


def _mail_th(text):
    style = ("border:1px solid #cccccc;font-family:Calibri,'맑은 고딕',sans-serif;"
             "font-size:10.5pt;color:#ffffff;background-color:#0d3b6e;"
             "padding:6px 10px;text-align:center;font-weight:bold;")
    return f'<td style="{style}">{text}</td>'


def build_email_html(report: "Report") -> str:
    """업무 메일 스타일 HTML 본문 — Table 기반 레이아웃 + Inline CSS (Outlook/모바일 호환).
    Summary(연한 파란색) / 삭제 완료 목록(연한 빨간색) / 수정 완료 목록(연한 초록색) /
    처리 실패 목록(연한 노란색, 있을 때만)을 순서대로 담는다."""
    s = report.summary
    today = datetime.date.today().strftime("%Y-%m-%d")
    FONT = "font-family:Calibri,'맑은 고딕',sans-serif;"
    BASE = FONT + "font-size:10.5pt;color:#333333;"

    # ── [작업 결과 요약] — 연한 파란색 ──
    summary_rows = "".join(
        f'<tr>{_mail_td(label)}{_mail_td(f"{val}건", "right", "font-weight:bold;")}</tr>'
        for label, val in [
            ("조회 부킹", s.get("total", 0)),
            ("ETD 지난 부킹", s.get("etd_over", 0)),
            ("삭제 완료", s.get("del_ok", 0)),
            ("수정 완료", s.get("mod_ok", 0)),
            ("처리 실패", s.get("fail", 0)),
        ]
    )
    summary_table = (
        '<table cellpadding="0" cellspacing="0" border="0" '
        'style="border-collapse:collapse;width:100%;background-color:#eef5fc;">'
        f"{summary_rows}</table>"
    )

    # ── [삭제 완료 부킹] — 연한 빨간색 ──
    deleted = [b for b in report.targets if b.action == "done_delete"]
    if deleted:
        rows = "".join(
            f"<tr>{_mail_td(i, 'center')}{_mail_td(b.booking_no)}"
            f"{_mail_td(str(b.etd) if b.etd else '-', 'center')}</tr>"
            for i, b in enumerate(deleted, start=1)
        )
        delete_block = (
            '<table cellpadding="0" cellspacing="0" border="0" '
            'style="border-collapse:collapse;width:100%;background-color:#fdecec;">'
            f"<tr>{_mail_th('No')}{_mail_th('Booking No')}{_mail_th('ETD')}</tr>"
            f"{rows}</table>"
        )
    else:
        delete_block = f'<p style="{BASE}margin:0;">삭제 완료된 부킹이 없습니다.</p>'

    # ── [수정 완료 부킹] — 연한 초록색 ──
    modified = [b for b in report.targets if b.action == "done_modify"]
    if modified:
        rows = "".join(
            f"<tr>{_mail_td(i, 'center')}{_mail_td(b.booking_no)}"
            f"{_mail_td(b.req_qty, 'center')}{_mail_td(b.new_qty, 'center')}"
            f"{_mail_td(str(b.etd) if b.etd else '-', 'center')}</tr>"
            for i, b in enumerate(modified, start=1)
        )
        modify_block = (
            '<table cellpadding="0" cellspacing="0" border="0" '
            'style="border-collapse:collapse;width:100%;background-color:#eaf7ea;">'
            f"<tr>{_mail_th('No')}{_mail_th('Booking No')}{_mail_th('기존 요청수량')}"
            f"{_mail_th('수정 후 수량')}{_mail_th('ETD')}</tr>"
            f"{rows}</table>"
        )
    else:
        modify_block = f'<p style="{BASE}margin:0;">수정 완료된 부킹이 없습니다.</p>'

    # ── [처리 실패] — 연한 노란색, 실패 건이 있을 때만 표시 ──
    failed = [b for b in report.targets if b.action in ("delete", "modify")]
    if failed:
        rows = "".join(
            f"<tr>{_mail_td(b.booking_no)}{_mail_td(b.fail_reason or '-')}</tr>"
            for b in failed
        )
        fail_block = (
            f'<h3 style="{BASE}color:#0d3b6e;margin:18px 0 6px 0;">[처리 실패]</h3>'
            '<table cellpadding="0" cellspacing="0" border="0" '
            'style="border-collapse:collapse;width:100%;background-color:#fff8e1;">'
            f"<tr>{_mail_th('Booking No')}{_mail_th('실패 사유')}</tr>"
            f"{rows}</table>"
        )
    else:
        fail_block = ""

    html = f"""<html>
<head><meta charset="utf-8"></head>
<body style="{BASE}margin:0;padding:0;">
<table cellpadding="0" cellspacing="0" border="0" style="width:100%;max-width:720px;{BASE}">
  <tr><td style="background-color:#0d3b6e;padding:16px 20px;">
    <span style="{FONT}font-size:14pt;color:#ffffff;font-weight:bold;">[BPT] 선반입 개수 관리 작업 결과 ({today})</span>
  </td></tr>
  <tr><td style="padding:18px 20px 6px 20px;">
    <p style="{BASE}margin:0 0 4px 0;">수신 : 담당자 제위</p>
    <p style="{BASE}margin:0 0 14px 0;">발신 : BPT 선반입 개수 관리 자동화 드림</p>
    <p style="{BASE}margin:0 0 4px 0;">안녕하십니까. &#128522;</p>
    <p style="{BASE}margin:0 0 4px 0;">항상 노고에 감사드립니다.</p>
    <p style="{BASE}margin:0 0 4px 0;">금일 BPT 선반입 개수 관리 자동화 작업이 정상적으로 완료되어 처리 결과를 안내드립니다.</p>
    <p style="{BASE}margin:0 0 16px 0;">상세 작업 내역은 첨부된 Excel 보고서를 참고 부탁드립니다.</p>

    <h3 style="{BASE}color:#0d3b6e;margin:0 0 6px 0;">[작업 결과 요약]</h3>
    {summary_table}

    <h3 style="{BASE}color:#0d3b6e;margin:18px 0 6px 0;">[삭제 완료 부킹]</h3>
    {delete_block}

    <h3 style="{BASE}color:#0d3b6e;margin:18px 0 6px 0;">[수정 완료 부킹]</h3>
    {modify_block}

    {fail_block}

    <h3 style="{BASE}color:#0d3b6e;margin:18px 0 6px 0;">[안내사항]</h3>
    <p style="{BASE}margin:0 0 4px 0;">&bull; 상세 작업 결과는 첨부된 Excel 보고서를 참고하여 주시기 바랍니다.</p>
    <p style="{BASE}margin:0 0 4px 0;">&bull; 삭제 또는 수정되지 않은 부킹은 ETD 여유 또는 처리 제외 대상입니다.</p>
    <p style="{BASE}margin:0 0 16px 0;">&bull; 본 메일은 BPT 선반입 개수 관리 자동화 프로그램에서 자동 생성되었습니다.</p>

    <p style="{BASE}margin:0;">감사합니다.<br>BPT 선반입 개수 관리 자동화 드림</p>
  </td></tr>
</table>
</body>
</html>
"""
    return html


def save_outlook_draft(report_path: Optional[str], html_body: str,
                        to_addr: str = REPORT_MAIL_TO, cc_addr: str = REPORT_MAIL_CC) -> bool:
    """Excel 보고서를 첨부해 Outlook 메일을 작성하고, 발송(Send) 없이 Draft(임시보관함)에만 저장한다.
    메일 창(Display)도 띄우지 않는다. 별도 워커 스레드에서 호출되므로 COM 아파트먼트를 스레드별로 초기화한다.
    cc_addr: 참조 — REPORT_MAIL_CC가 비어있으면 설정하지 않음 (향후 확장용)."""
    if not report_path or not os.path.exists(report_path):
        print(f"[메일] 첨부파일이 존재하지 않아 메일을 생성하지 않음: {report_path}")
        return False

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as e:
            print(f"[메일] Outlook 실행 실패: {e}")
            return False

        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = to_addr
        if cc_addr:
            mail.CC = cc_addr
        mail.Subject = f"[BPT] 선반입 개수 관리 작업 결과 ({datetime.date.today()})"
        mail.HTMLBody = html_body
        mail.Attachments.Add(report_path)
        mail.Save()   # ※ Send() 사용 금지 — Drafts(임시보관함)에만 저장
        return True
    except Exception as e:
        print(f"[메일] Draft 저장 실패: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()


def send_report_mail(all_bookings, targets, log_func=print) -> bool:
    """삭제/수정 작업이 모두 끝난 마지막 단계에서 호출.
    ⑤ Excel 보고서 생성 → ⑥ Outlook 메일(HTML) 작성 → ⑦ Draft(임시보관함) 저장까지 수행한다.
    실제 발송은 하지 않으며, 각 단계 실패는 로그만 남기고 프로그램은 계속 동작한다."""
    report = generate_report(all_bookings, targets)

    report_path = export_excel_report(report)
    if not report_path:
        log_func("Excel 보고서 생성 실패")
        return False

    html = build_email_html(report)

    if not save_outlook_draft(report_path, html):
        log_func("Outlook 메일 작성/Draft 저장 실패")
        return False

    filename = os.path.basename(report_path)
    log_func(
        "==============================\n"
        "Excel 보고서 생성 완료\n"
        "Outlook 메일 작성 완료\n"
        "Draft 저장 완료\n"
        f"받는 사람 {REPORT_MAIL_TO}\n"
        f"첨부파일 {filename}\n"
        "=============================="
    )
    return True


# ══════════════════════════════════════════════════════
# 메인 앱
# ══════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BPT 부킹 선반입 관리")
        self.geometry("1100x750")
        self.configure(bg="#0a0f1a")
        self.resizable(True, True)

        self.bookings: list[Booking] = []
        self.log_q = queue.Queue()
        self._running = False
        self._browser = None
        self._page = None
        self._pw = None
        self._pw_instance = None

        self._build_ui()
        self._poll_log()

    # ── UI 구성 ──────────────────────────────────────
    def _build_ui(self):
        # 제목
        hdr = tk.Frame(self, bg="#0a0f1a")
        hdr.pack(fill="x", padx=20, pady=(18, 0))
        tk.Label(hdr, text="⚓  BPT 부킹 선반입 관리 자동화",
                 font=("맑은 고딕", 16, "bold"),
                 fg="#7ec8f0", bg="#0a0f1a").pack(side="left")
        tk.Label(hdr, text=f"기준일: {datetime.date.today()} | 내일: {get_tomorrow()}",
                 font=("맑은 고딕", 9), fg="#4a6a8a", bg="#0a0f1a").pack(side="right")

        # 버튼 행
        btn_frame = tk.Frame(self, bg="#0a0f1a")
        btn_frame.pack(fill="x", padx=20, pady=10)

        self.btn_fetch = self._btn(btn_frame, "① 선반입 부킹 조회", self._start_fetch, "#0d3b6e", "#7ec8f0")
        self.btn_fetch.pack(side="left", padx=(0, 8))

        self.btn_etd_all = self._btn(btn_frame, "② ETD 확인 (전체)", self._start_etd_all, "#1a3a1a", "#7ef07e")
        self.btn_etd_all.pack(side="left", padx=(0, 8))
        self.btn_etd_all.config(state="disabled")

        self.btn_etd = self._btn(btn_frame, "③ ETD 확인 (선택)", self._start_etd_selected, "#1a2a3a", "#7ec8f0")
        self.btn_etd.pack(side="left", padx=(0, 8))
        self.btn_etd.config(state="disabled")

        self.btn_apply = self._btn(btn_frame, "④ 선택 항목 처리", self._apply_selected, "#3a1a0d", "#f0c07e")
        self.btn_apply.pack(side="left", padx=(0, 8))
        self.btn_apply.config(state="disabled")

        self.btn_apply_all = self._btn(btn_frame, "⑤ 전체 처리", self._apply_all, "#2a0d0d", "#f07e7e")
        self.btn_apply_all.pack(side="left", padx=(0, 8))
        self.btn_apply_all.config(state="disabled")

        # v4: 보고서를 Outlook 임시보관함(Draft)에 수동으로 저장하는 버튼
        self.btn_mail_draft = self._btn(btn_frame, "⑥ 보고서 Outlook 임시보관함 저장",
                                         self._start_report_mail, "#1a1a3a", "#c0a8f0")
        self.btn_mail_draft.pack(side="left", padx=(0, 8))
        self.btn_mail_draft.config(state="disabled")

        tk.Label(btn_frame, text="⚠ Sinokor Portal 열어두세요",
                 font=("맑은 고딕", 9), fg="#f0b07e", bg="#0a0f1a").pack(side="right")

        # 테이블
        tbl_frame = tk.Frame(self, bg="#0a0f1a")
        tbl_frame.pack(fill="both", expand=True, padx=20, pady=(0, 8))

        cols = ("No", "선택", "부킹번호", "요청수량", "반입수량", "확인상태", "ETD", "처리내용")
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", height=16)

        widths = [40, 50, 180, 80, 80, 100, 110, 180]
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")

        sb_y = ttk.Scrollbar(tbl_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb_y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb_y.pack(side="right", fill="y")

        self.tree.bind("<Button-1>", self._on_click)
        self.tree.bind("<Double-Button-1>", self._copy_booking_no)
        self.tree.bind("<Control-c>", self._copy_booking_no)
        self.tree.column("No", width=40, anchor="center")
        self.tree.heading("선택", command=self._toggle_all)
        self._all_selected = False

        # 태그 색상
        self.tree.tag_configure("modify", background="#1a2a1a", foreground="#7ef07e")
        self.tree.tag_configure("delete", background="#2a1a1a", foreground="#f07e7e")
        self.tree.tag_configure("skip",   background="#0f0f1a", foreground="#4a6a8a")
        self.tree.tag_configure("pending",background="#0a0f1a", foreground="#cccccc")

        # 스타일
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#0d1220", foreground="#cccccc",
                        fieldbackground="#0d1220", rowheight=26,
                        font=("맑은 고딕", 9))
        style.configure("Treeview.Heading", background="#0d3b6e", foreground="#7ec8f0",
                        font=("맑은 고딕", 9, "bold"))
        style.map("Treeview", background=[("selected", "#1a4a7a")])

        # 로그창
        log_frame = tk.Frame(self, bg="#0a0f1a")
        log_frame.pack(fill="x", padx=20, pady=(0, 12))
        tk.Label(log_frame, text="로그", font=("맑은 고딕", 8),
                 fg="#4a6a8a", bg="#0a0f1a").pack(anchor="w")
        self.log_box = scrolledtext.ScrolledText(
            log_frame, height=8, bg="#060a12", fg="#5a8a5a",
            font=("Consolas", 8), state="disabled", bd=0)
        self.log_box.pack(fill="x")

    def _btn(self, parent, text, cmd, bg, fg):
        b = tk.Button(parent, text=text, command=cmd,
                      bg=bg, fg=fg, activebackground=bg,
                      font=("맑은 고딕", 10, "bold"),
                      relief="flat", padx=16, pady=8, cursor="hand2")
        b.bind("<Enter>", lambda e: b.config(bg=self._lighten(bg)))
        b.bind("<Leave>", lambda e: b.config(bg=bg))
        return b

    def _lighten(self, hex_color):
        r, g, b = int(hex_color[1:3],16), int(hex_color[3:5],16), int(hex_color[5:7],16)
        r, g, b = min(255,r+30), min(255,g+30), min(255,b+30)
        return f"#{r:02x}{g:02x}{b:02x}"

    # ── 로그 ──────────────────────────────────────────
    def log(self, msg):
        self.log_q.put(msg)

    def _poll_log(self):
        while not self.log_q.empty():
            msg = self.log_q.get()
            self.log_box.config(state="normal")
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            self.log_box.insert("end", f"[{ts}] {msg}\n")
            self.log_box.see("end")
            self.log_box.config(state="disabled")
        self.after(100, self._poll_log)

    # ── 테이블 클릭 (선택 토글) ──────────────────────
    def _copy_booking_no(self, event=None):
        """선택된 행의 부킹번호 클립보드 복사"""
        # 더블클릭 시 해당 행 직접 찾기
        if event and event.type == '4':  # Button event
            item = self.tree.identify_row(event.y)
        else:
            selected = self.tree.selection()
            item = selected[0] if selected else None

        if item:
            vals = self.tree.item(item, "values")
            if len(vals) > 2:
                bno = vals[2]
                self.clipboard_clear()
                self.clipboard_append(bno)
                self.log(f"복사됨: {bno}")

    def _toggle_all(self):
        """선택 헤더 클릭 시 전체 선택/해제"""
        self._all_selected = not self._all_selected
        for item in self.tree.get_children():
            vals = list(self.tree.item(item, "values"))
            vals[1] = "☑" if self._all_selected else "☐"
            self.tree.item(item, values=vals)

    def _auto_check(self):
        """요청≠반입 또는 삭제 상태인 항목 자동 체크"""
        for item in self.tree.get_children():
            vals = list(self.tree.item(item, "values"))
            bno = vals[2]
            bk = next((b for b in self.bookings if b.booking_no == bno), None)
            if bk is None:
                continue
            if bk.recv_qty is None or bk.req_qty != bk.recv_qty:
                vals[1] = "☑"
            else:
                vals[1] = "☐"
            self.tree.item(item, values=vals)

    def _on_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        col    = self.tree.identify_column(event.x)
        if region == "cell" and col == "#2":
            item = self.tree.identify_row(event.y)
            if item:
                vals = list(self.tree.item(item, "values"))
                vals[1] = "☑" if vals[1] == "☐" else "☐"
                self.tree.item(item, values=vals)

    def _refresh_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for idx, bk in enumerate(self.bookings, start=1):
            # 요청≠반입이면 자동 체크
            if bk.action in ("modify", "delete"):
                checked = "☑"
            elif bk.action in ("same", "skip"):
                checked = "☐"
            elif bk.recv_qty is None or bk.req_qty != bk.recv_qty:
                checked = "☑"
            else:
                checked = "☐"
            etd_str = str(bk.etd) if bk.etd else "-"
            if bk.action == "modify":
                action_str = f"수정 → {bk.new_qty} van"
                tag = "modify"
            elif bk.action == "delete":
                action_str = "삭제"
                tag = "delete"
            elif bk.action == "same":
                action_str = ""
                tag = "skip"
                checked = "☐"
            elif bk.action == "skip":
                action_str = "스킵 (ETD 여유)"
                tag = "skip"
                checked = "☐"
            else:
                action_str = "ETD 미확인"
                tag = "pending"
                checked = "☐"
            self.tree.insert("", "end", values=(
                idx, checked, bk.booking_no, bk.req_qty,
                bk.recv_qty or "-", bk.confirm,
                etd_str, action_str
            ), tags=(tag,))

    # ── ① BPT 조회 ────────────────────────────────────
    def _start_fetch(self):
        if self._running:
            return
        self._running = True
        self.btn_fetch.config(state="disabled", text="조회 중...")
        threading.Thread(target=self._fetch_thread, daemon=True).start()

    def _fetch_thread(self):
        try:
            self.log("BPT 사이트 접속 중...")
            from playwright.sync_api import sync_playwright
            pw = sync_playwright().start()
            self._pw_instance = pw
            if True:
                # exe 환경에서 시스템 Chrome 사용
                import os as _os
                # 현재 사용자 AppData 경로 동적 포함
                _user = _os.environ.get("USERNAME","")
                chrome_paths = [
                    'C:/Program Files/Google/Chrome/Application/chrome.exe',
                    'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
                    f'C:/Users/{_user}/AppData/Local/Google/Chrome/Application/chrome.exe',
                ]
                chrome_exe = next((p for p in chrome_paths if _os.path.exists(p)), None)
                if chrome_exe:
                    browser = pw.chromium.launch(headless=False, executable_path=chrome_exe)
                else:
                    browser = pw.chromium.launch(headless=False)
                page = browser.new_page()
                page.goto(BPT_URL, wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(2000)

                # 팝업 닫기
                for sel in ["text=닫기","text=오늘 하루동안 보지 않기","a:has-text('닫기')","button:has-text('닫기')"]:
                    try:
                        btns = page.locator(sel)
                        for i in range(btns.count()):
                            if btns.nth(i).is_visible(timeout=500):
                                btns.nth(i).click()
                                page.wait_for_timeout(600)
                    except Exception:
                        pass

                # 로그인
                page.click("text=로그인 & 회원가입")
                page.wait_for_timeout(2000)
                for sel in ["input[name='userId']","input[type='text']"]:
                    try:
                        loc = page.locator(sel).first
                        if loc.is_visible(timeout=2000):
                            loc.fill(BPT_ID); break
                    except Exception: continue
                for sel in ["input[type='password']"]:
                    try:
                        loc = page.locator(sel).first
                        if loc.is_visible(timeout=2000):
                            loc.fill(BPT_PW); break
                    except Exception: continue
                try:
                    btn = page.locator("button:has-text('로그인'),input[value='로그인']").first
                    if btn.is_visible(timeout=2000): btn.click()
                    else: page.keyboard.press("Enter")
                except Exception:
                    page.keyboard.press("Enter")
                page.wait_for_load_state("domcontentloaded", timeout=20000)
                page.wait_for_timeout(1000)
                self.log("로그인 완료")

                # 부킹 선반입 관리 이동
                page.goto(BOOKING_URL, wait_until="domcontentloaded", timeout=20000)
                page.wait_for_timeout(2000)

                # 선사코드 조회
                all_frames = [page] + list(page.frames)
                for frame in all_frames:
                    for sel in ["input[name='crrCd']","table input[type='text']","input[type='text']"]:
                        try:
                            loc = frame.locator(sel).first
                            if loc.is_visible(timeout=1000):
                                loc.fill(CARRIER); break
                        except Exception: continue
                    else: continue
                    break

                for frame in all_frames:
                    for sel in ["input[value='조회']","button:has-text('조회')"]:
                        try:
                            loc = frame.locator(sel).first
                            if loc.is_visible(timeout=1000):
                                loc.click()
                                page.wait_for_timeout(3000)
                                break
                        except Exception: continue
                    else: continue
                    break

                # 파싱
                rows_all = []
                for frame in all_frames:
                    try:
                        r = frame.query_selector_all("table tr")
                        if len(r) > len(rows_all): rows_all = r
                    except Exception: continue

                bookings = []
                for row in rows_all:
                    cells = row.query_selector_all("td")
                    if len(cells) < 6: continue
                    texts = [c.inner_text().strip() for c in cells]
                    bno = texts[0]
                    if not re.match(r'^[A-Z]{4}\d{10,}', bno): continue
                    try: req = int(texts[4]) if texts[4] else 0
                    except: req = 0
                    try: recv = int(texts[5]) if texts[5] else None
                    except: recv = None
                    confirm = texts[6] if len(texts) > 6 else ""
                    bookings.append(Booking(bno, req, recv, confirm))

                # 브라우저 유지 (처리 시 재사용)
                self._browser = browser
                self._page = page
                # playwright 인스턴스도 유지 (닫히지 않게)

            self.bookings = bookings
            self.log(f"총 {len(bookings)}건 수집 완료")
            self.after(0, self._refresh_table)
            self.after(0, self._auto_check)
            # 선택 건수 로그
            checked = sum(1 for b in bookings if b.recv_qty is None or b.req_qty != b.recv_qty)
            self.log(f"총 {len(bookings)}건 수집 완료 / {checked}건 선택됨")
            self.after(0, lambda: self.btn_etd_all.config(state="normal"))
            self.after(0, lambda: self.btn_etd.config(state="normal"))
            self.after(0, lambda: self.btn_mail_draft.config(state="normal"))  # v4

        except Exception as e:
            self.log(f"오류: {e}")
        finally:
            self._running = False
            self.after(0, lambda: self.btn_fetch.config(state="normal", text="① BPT 조회"))

    # ── ② ETD 확인 ────────────────────────────────────
    def _start_etd_all(self):
        """전체 ETD 확인"""
        if self._running:
            return
        if not messagebox.askyesno("ETD 확인",
            "Sinokor Portal V2가 열려있고\nBooking Management 화면이 활성화되어 있나요?"):
            return
        self._running = True
        self.btn_etd_all.config(state="disabled", text="ETD 확인 중...")
        threading.Thread(target=self._etd_thread, args=(None,), daemon=True).start()

    def _start_etd_selected(self):
        """선택된 항목만 ETD 확인"""
        if self._running:
            return
        # 체크된 부킹번호 수집
        checked_bnos = set()
        for item in self.tree.get_children():
            vals = self.tree.item(item, "values")
            if vals[1] == "☑":
                checked_bnos.add(vals[2])
        if not checked_bnos:
            messagebox.showwarning("선택 없음", "☑ 체크된 항목이 없습니다.")
            return
        if not messagebox.askyesno("ETD 확인",
            f"선택된 {len(checked_bnos)}건만 ETD 확인합니다.\nSinokor Portal이 열려있나요?"):
            return
        self._running = True
        self.btn_etd.config(state="disabled", text="ETD 확인 중...")
        threading.Thread(target=self._etd_thread, args=(checked_bnos,), daemon=True).start()

    def _start_etd(self):
        pass  # 하위 호환용

    def _etd_thread(self, filter_bnos=None):
        tomorrow = get_tomorrow()
        try:
            # 캐시 로드
            etd_cache = load_etd_cache()
            cache_hits = sum(1 for bk in self.bookings if bk.booking_no in etd_cache)
            if etd_cache:
                self.log(f"오늘 캐시 발견 — {cache_hits}건 재사용, 나머지만 Sinokor 조회")
            else:
                self.log("캐시 없음 — 전체 Sinokor 조회 시작")

            for i, bk in enumerate(self.bookings):
                # 선택 필터 적용 (선택 모드일 때)
                if filter_bnos is not None and bk.booking_no not in filter_bnos:
                    continue

                # 요청수량 == 반입수량이면 ETD 확인 없이 스킵
                if bk.recv_qty and bk.req_qty == bk.recv_qty:
                    bk.action = "same"
                    self.log(f"[{i+1}/{len(self.bookings)}] {bk.booking_no} 스킵 (요청=반입={bk.req_qty})")
                    continue

                # 캐시에 있으면 재사용 (None 이면 재조회)
                if bk.booking_no in etd_cache and etd_cache[bk.booking_no] is not None:
                    cached = etd_cache[bk.booking_no]
                    bk.etd = parse_etd(cached)
                    self.log(f"[{i+1}/{len(self.bookings)}] {bk.booking_no} 캐시 사용: {bk.etd}")
                    # action 결정
                    tomorrow = get_tomorrow()
                    if bk.etd is None:
                        bk.action = ""
                    elif bk.etd > tomorrow:
                        bk.action = "skip"
                    elif bk.recv_qty and bk.recv_qty > 0:
                        bk.action = "modify"
                        bk.new_qty = bk.recv_qty
                    else:
                        bk.action = "delete"
                    continue
                elif bk.booking_no in etd_cache and etd_cache[bk.booking_no] is None:
                    self.log(f"[{i+1}/{len(self.bookings)}] {bk.booking_no} 캐시 None — Portal 재조회")

                self.log(f"[{i+1}/{len(self.bookings)}] {bk.booking_no} ETD 확인 중...")
                etd = self._get_etd(bk.booking_no)
                bk.etd = etd
                # 캐시에 저장
                etd_cache[bk.booking_no] = str(etd) if etd else None

                if etd is None:
                    bk.action = ""
                elif etd > tomorrow:
                    bk.action = "skip"
                elif bk.recv_qty and bk.recv_qty > 0:
                    if bk.req_qty == bk.recv_qty:
                        bk.action = "same"
                    else:
                        bk.action = "modify"
                        bk.new_qty = bk.recv_qty
                else:
                    bk.action = "delete"

            # 캐시 파일 저장
            save_etd_cache(etd_cache)
            self.log("ETD 캐시 저장 완료")

            self.after(0, self._refresh_table)
            n = sum(1 for b in self.bookings if b.action in ("modify","delete"))
            self.log(f"ETD 확인 완료 — 처리 대상 {n}건")
            self.after(0, lambda: self.btn_etd_all.config(state="normal"))
            self.after(0, lambda: self.btn_etd.config(state="normal"))
            self.after(0, lambda: self.btn_apply.config(state="normal"))
            self.after(0, lambda: self.btn_apply_all.config(state="normal"))

        except Exception as e:
            self.log(f"ETD 확인 오류: {e}")
        finally:
            self._running = False
            self.after(0, lambda: self.btn_etd_all.config(state="normal", text="② ETD 확인 (전체)"))
            self.after(0, lambda: self.btn_etd.config(state="normal", text="③ ETD 확인 (선택)"))
            self.after(0, lambda: self.btn_apply.config(state="normal"))
            self.after(0, lambda: self.btn_apply_all.config(state="normal"))

    def _find_portal_window(self):
        """Sinokor Portal 창 탐색 (대소문자 무시)"""
        keywords = ["skrportal", "portal system", "sinokor", "시노코르", "sino", "portal v2"]
        desktop = Desktop(backend="uia")
        windows = desktop.windows()
        for w in windows:
            t = w.window_text()
            if any(k in t.lower() for k in keywords):
                return w
        titles = [w.window_text() for w in windows if w.window_text().strip()]
        self.log(f"  ⚠ Portal 쉕 못 찾음. 열린 쉕목록: {titles[:10]}")
        return None

    def _close_popups(self):
        """알림/경고/안내 팝업 자동 닫기.
        포탈에서 뜨는 작은 다이얼로그(OK 버튼 1개)만 닫는다.
        Chrome/브라우저 창은 절대 건드리지 않음.
        """
        # 절대 건드리지 않을 창 키워드
        EXCLUDE = ["Chrome","Firefox","Edge","BPT","Portal System","SKRPORTAL","Whale","Opera","Naver"]
        # 팝업으로 인식할 창 제목 키워드
        POPUP_TITLES = [
            "경고","알림","오류","없습니다","Alert","Error","Warning",
            "Booking Management",
            "Information",
            "Confirm",   # "Confirm Shipping Status" 팝업
        ]
        # OK/확인 계열 → 클릭, No/취소 계열 → 클릭 (Ship Cfm 팝업은 No로 닫아야 함)
        OK_NAMES = ["OK","Ok","확인","ok","닫기","Close","No","NO","취소"]

        def find_ok(node, depth=0):
            if depth > 3:
                return None
            try:
                n = node.element_info.name or ""
                ct = getattr(node.element_info, "control_type", "") or ""
                if ct == "Button" and any(n == ok for ok in OK_NAMES):
                    return node
                for child in node.children():
                    result = find_ok(child, depth + 1)
                    if result:
                        return result
            except Exception:
                pass
            return None

        try:
            desktop = Desktop(backend="uia")
            for w in desktop.windows():
                t = w.window_text()
                # 브라우저/대형 앱 창 제외
                if any(k in t for k in EXCLUDE):
                    continue
                # 팝업 키워드 매칭
                if not any(k in t for k in POPUP_TITLES):
                    continue
                ok_btn = find_ok(w)
                if ok_btn:
                    ok_btn.click_input()
                    time.sleep(0.4)
                    self.log(f"  팝업 닫음 (OK클릭): {t}")
                    return
                else:
                    pyautogui.press("enter")
                    time.sleep(0.4)
                    self.log(f"  팝업 닫음 (Enter): {t}")
                    return
        except Exception:
            pass

    def _get_etd(self, booking_no: str) -> Optional[datetime.date]:
        try:
            win = self._find_portal_window()
            if win is None:
                self.log("  Sinokor Portal 창 못 찾음")
                return None

            win.set_focus()
            time.sleep(0.5)

            # ── Booking Management 탭 내 컨트롤 탐색 ──
            bkg_win = None
            try:
                for child in win.children():
                    ei = child.element_info
                    aid_val = getattr(ei, "automation_id", None) or getattr(ei, "auto_id", None) or ""
                    if str(aid_val) == "329602":       # MDICLIENT Pane
                        for sub in child.children():
                            sub_name = sub.element_info.name or ""
                            if "Booking Management" in sub_name:
                                bkg_win = sub
                                break
                        break
            except Exception:
                pass

            # bkg_win 못 찾으면 win 전체에서 탐색
            search_root = bkg_win if bkg_win else win

            def get_aid(elem):
                ei = elem.element_info
                for attr in ("automation_id", "auto_id"):
                    try:
                        v = getattr(ei, attr, None)
                        if v is not None: return str(v)
                    except Exception: pass
                return ""

            def get_name(elem):
                try: return elem.element_info.name or ""
                except Exception: return ""

            def get_ct(elem):
                try: return elem.element_info.control_type or ""
                except Exception: return ""

            def find_by_aid(root, aid, max_depth=6):
                """automation_id로 컨트롤 재귀 탐색"""
                try:
                    if get_aid(root) == aid:
                        return root
                    if max_depth <= 0:
                        return None
                    for child in root.children():
                        result = find_by_aid(child, aid, max_depth - 1)
                        if result:
                            return result
                except Exception:
                    pass
                return None

            def find_by_name(root, name, ctrl_type=None, max_depth=5):
                """name으로 컨트롤 재귀 탐색"""
                try:
                    n = get_name(root)
                    ct = get_ct(root)
                    if n == name and (ctrl_type is None or ct == ctrl_type):
                        return root
                    if max_depth <= 0:
                        return None
                    for child in root.children():
                        result = find_by_name(child, name, ctrl_type, max_depth - 1)
                        if result:
                            return result
                except Exception:
                    pass
                return None

            # ── ① B/K No. 입력 (auto_id='BK_NO' ComboBox 내부 Edit) ──
            bkno_ctrl = None
            bkno_combo = find_by_aid(search_root, "BK_NO")
            if bkno_combo:
                # ComboBox 내부 Edit (auto_id='1001')
                try:
                    for child in bkno_combo.children():
                        ct = child.element_info.control_type or ""
                        if ct == "Edit":
                            bkno_ctrl = child
                            break
                except Exception:
                    pass
                if bkno_ctrl is None:
                    bkno_ctrl = bkno_combo  # fallback

            if bkno_ctrl is None:
                self.log("  ⚠ B/K No. 컨트롤 못 찾음 — 좌표 fallback (다른 PC에서는 오작동 가능)")
                win.set_focus(); time.sleep(0.3)
                pyautogui.click(107, 136); time.sleep(0.3)
                pyautogui.hotkey("ctrl","a")
                pyperclip.copy(booking_no)
                pyautogui.hotkey("ctrl","v"); time.sleep(0.3)
            else:
                bkno_ctrl.click_input()
                time.sleep(0.3)
                # set_edit_text로 직접 값 설정 시도 (키보드 시뮬레이션 불필요)
                _input_ok = False
                try:
                    bkno_ctrl.set_edit_text(booking_no)
                    time.sleep(0.2)
                    _input_ok = True
                    self.log(f"  B/K No. 입력(set_edit_text): {booking_no}")
                except Exception:
                    pass
                # fallback: pyautogui 로 Ctrl+A → Ctrl+V
                if not _input_ok:
                    pyperclip.copy(booking_no)
                    pyautogui.hotkey("ctrl", "a")
                    time.sleep(0.15)
                    pyautogui.hotkey("ctrl", "v")
                    time.sleep(0.3)
                    self.log(f"  B/K No. 입력(hotkey): {booking_no}")

            # ── ② Search 전 — Ship Cfm 팝업 방지: B/K No. 필드로 포커스 고정 ──
            # (이전 부킹 화면에서 Ship Cfm 체크박스에 포커스가 남아있으면
            #  Search 클릭 시 Confirm Shipping Status 팝업이 뜰 수 있음)
            if bkno_ctrl:
                try:
                    bkno_ctrl.click_input()
                    time.sleep(0.1)
                except Exception:
                    pass

            # ── ② Search 버튼 클릭 (auto_id='btnSearch') ──
            search_btn = find_by_aid(win, "btnSearch", max_depth=8)
            if search_btn is None:
                for sname in ["Se&arch", "Search", "&Search"]:
                    search_btn = find_by_name(win, sname, ctrl_type="Button", max_depth=8)
                    if search_btn:
                        break
            if search_btn is None:
                self.log("  ⚠ Search 버튼 못 찾음 — 좌표 fallback (다른 PC에서는 오작동 가능)")
                win.set_focus(); time.sleep(0.3)
                pyautogui.click(592, 105); time.sleep(3.0)
            else:
                search_btn.click_input()
                self.log("  Search 클릭 (btnSearch)")
                time.sleep(3.0)

            # ── 팝업 처리 (Ship Cfm 확인창 포함) ──
            self._close_popups()
            time.sleep(0.3)
            # 혹시 팝업이 더 있을 수 있으므로 한번 더
            self._close_popups()
            time.sleep(0.3)

            # ── ③ ETD 필드 읽기 (auto_id='TML_ETD') ──
            # search_root 갱신 (Search 후 bkg_win 재탐색)
            try:
                for child in win.children():
                    ei = child.element_info
                    aid_val = getattr(ei, "automation_id", None) or getattr(ei, "auto_id", None) or ""
                    if str(aid_val) == "329602":
                        for sub in child.children():
                            sub_name = sub.element_info.name or ""
                            if "Booking Management" in sub_name:
                                search_root = sub
                                break
                        break
            except Exception:
                pass

            # ── gcSchedule > ETD (Schedule Information ETD) ──
            # TML_ETD(터미널 ETD)가 아닌 gcSchedule 패널의 ETD를 읽어야 함
            gc_schedule = find_by_aid(search_root, "gcSchedule")
            if gc_schedule is None:
                gc_schedule = find_by_aid(win, "gcSchedule", max_depth=8)

            etd_ctrl = None
            if gc_schedule:
                etd_ctrl = find_by_aid(gc_schedule, "ETD", max_depth=3)

            etd_text = ""
            if etd_ctrl is None:
                self.log("  ⚠ ETD 컨트롤(gcSchedule>ETD) 못 찾음 — 직접 탐색 시도")
                for aid_candidate in ("TML_ETD", "ETD", "etd"):
                    etd_ctrl = find_by_aid(search_root, aid_candidate, max_depth=8)
                    if etd_ctrl is None:
                        etd_ctrl = find_by_aid(win, aid_candidate, max_depth=8)
                    if etd_ctrl:
                        self.log(f"  ETD 컨트롤 발견 (aid={aid_candidate})")
                        break
                if etd_ctrl is None:
                    self.log("  ETD 컨트롤 완전히 못 찾음")
            else:
                # 방법1: get_value (가장 정확)
                try:
                    etd_text = etd_ctrl.get_value().strip()
                except Exception:
                    pass
                # 방법2: legacy_properties Value
                if not etd_text:
                    try:
                        etd_text = etd_ctrl.legacy_properties().get("Value", "").strip()
                    except Exception:
                        pass
                # 방법3: 클립보드
                if not etd_text:
                    try:
                        pyperclip.copy("")
                        etd_ctrl.click_input()
                        time.sleep(0.2)
                        etd_ctrl.type_keys("^a^c", with_spaces=False)
                        time.sleep(0.3)
                        etd_text = pyperclip.paste().strip()
                    except Exception:
                        pass
                self.log(f"  ETD raw: '{etd_text}'")

            etd = parse_etd(etd_text)
            if etd:
                self.log(f"  {booking_no} ETD: {etd}")
            else:
                self.log(f"  {booking_no} ETD 파싱 실패: '{etd_text}'")
            return etd

        except Exception as e:
            self.log(f"  ETD 오류: {e}")
            return None

    # ── ③ 처리 ────────────────────────────────────────
    def _apply_selected(self):
        items = self.tree.get_children()
        selected = []
        for item in items:
            vals = self.tree.item(item, "values")
            if vals[1] == "☑":
                bno = vals[2]
                bk = next((b for b in self.bookings if b.booking_no == bno), None)
                if bk and bk.action in ("modify","delete"):
                    selected.append(bk)
        if not selected:
            messagebox.showwarning("선택 없음", "☑ 체크된 처리 대상 항목이 없습니다.")
            return
        if not messagebox.askyesno("확인", f"{len(selected)}건을 처리하시겠습니까?"):
            return
        self._start_process(selected)

    def _apply_all(self):
        targets = [b for b in self.bookings if b.action in ("modify","delete")]
        if not targets:
            messagebox.showinfo("없음", "처리할 항목이 없습니다.")
            return
        msg = f"전체 {len(targets)}건을 처리하시겠습니까?\n\n"
        msg += "\n".join(f"  {'수정' if b.action=='modify' else '삭제'}: {b.booking_no}" for b in targets[:10])
        if len(targets) > 10:
            msg += f"\n  ... 외 {len(targets)-10}건"
        if not messagebox.askyesno("전체 처리 확인", msg):
            return
        self._start_process(targets)

    def _start_process(self, targets):
        if self._running:
            return
        self._running = True
        self.btn_apply.config(state="disabled")
        self.btn_apply_all.config(state="disabled")
        threading.Thread(target=self._process_thread, args=(targets,), daemon=True).start()

    def _process_thread(self, targets):
        try:
            self.log(f"처리 시작 ({len(targets)}건)...")
            from playwright.sync_api import sync_playwright
            import os as _os

            _user = _os.environ.get("USERNAME","")
            chrome_paths = [
                'C:/Program Files/Google/Chrome/Application/chrome.exe',
                'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
                f'C:/Users/{_user}/AppData/Local/Google/Chrome/Application/chrome.exe',
            ]
            chrome_exe = next((p for p in chrome_paths if _os.path.exists(p)), None)

            with sync_playwright() as pw:
                if chrome_exe:
                    browser = pw.chromium.launch(headless=False, executable_path=chrome_exe)
                else:
                    browser = pw.chromium.launch(headless=False)
                page = browser.new_page()

                # 로그인
                page.goto(BPT_URL, wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(2000)
                for sel in ["text=닫기","a:has-text('닫기')","button:has-text('닫기')"]:
                    try:
                        btns = page.locator(sel)
                        for i in range(btns.count()):
                            if btns.nth(i).is_visible(timeout=500):
                                btns.nth(i).click(); page.wait_for_timeout(500)
                    except Exception: pass

                page.click("text=로그인 & 회원가입")
                page.wait_for_timeout(2000)
                for sel in ["input[name='userId']","input[type='text']"]:
                    try:
                        loc = page.locator(sel).first
                        if loc.is_visible(timeout=2000): loc.fill(BPT_ID); break
                    except Exception: continue
                for sel in ["input[type='password']"]:
                    try:
                        loc = page.locator(sel).first
                        if loc.is_visible(timeout=2000): loc.fill(BPT_PW); break
                    except Exception: continue
                try:
                    btn = page.locator("button:has-text('로그인'),input[value='로그인']").first
                    if btn.is_visible(timeout=2000): btn.click()
                    else: page.keyboard.press("Enter")
                except Exception:
                    page.keyboard.press("Enter")
                page.wait_for_load_state("domcontentloaded", timeout=20000)
                page.wait_for_timeout(1000)
                self.log("로그인 완료")

                # 부킹 선반입 관리
                page.goto(BOOKING_URL, wait_until="domcontentloaded", timeout=20000)
                page.wait_for_timeout(2000)
                all_frames = [page] + list(page.frames)
                for frame in all_frames:
                    for sel in ["input[name='crrCd']","table input[type='text']","input[type='text']"]:
                        try:
                            loc = frame.locator(sel).first
                            if loc.is_visible(timeout=1000): loc.fill(CARRIER); break
                        except Exception: continue
                    else: continue
                    break
                for frame in all_frames:
                    for sel in ["input[value='조회']","button:has-text('조회')"]:
                        try:
                            loc = frame.locator(sel).first
                            if loc.is_visible(timeout=1000):
                                loc.click(); page.wait_for_timeout(3000); break
                        except Exception: continue
                    else: continue
                    break

                # 각 항목 처리 — iframe 포함 전체 탐색
                all_frames = [page] + list(page.frames)

                # 가장 많은 tr을 가진 frame을 active_frame으로
                active_frame = page
                max_rows = 0
                for frame in all_frames:
                    try:
                        cnt = frame.locator("table tr").count()
                        if cnt > max_rows:
                            max_rows = cnt
                            active_frame = frame
                    except Exception:
                        continue
                self.log(f"  활성 프레임: rows={max_rows}")

                for bk in targets:
                    # 행 찾기
                    row_loc = None
                    for frame in [active_frame] + [f for f in all_frames if f != active_frame]:
                        rl = frame.locator(f"tr:has-text('{bk.booking_no}')")
                        if rl.count():
                            row_loc = rl
                            active_frame = frame
                            break

                    if row_loc is None:
                        self.log(f"  {bk.booking_no} 행 없음, 스킵")
                        bk.fail_reason = "부킹 행을 찾을 수 없음"  # v4: 보고서용 실패 사유
                        continue

                    self.log(f"  {bk.booking_no} 행 발견")

                    if bk.action == "delete":
                        done = False
                        for sel in [
                            "a:has-text('삭제')",
                            "a[onclick*='Del']",
                            "a[onclick*='del']",
                            "span.btn_pack6 a",
                            "input[value='삭제']",
                            "button:has-text('삭제')",
                        ]:
                            try:
                                btn = row_loc.locator(sel).first
                                if btn.count() and btn.is_visible(timeout=1000):
                                    def handle_dialog(dialog):
                                        dialog.accept()
                                    page.once("dialog", handle_dialog)
                                    btn.click()
                                    page.wait_for_timeout(2000)
                                    self.log(f"  {bk.booking_no} 삭제 완료")
                                    bk.action = "done_delete"
                                    done = True
                                    break
                            except Exception as e:
                                self.log(f"  삭제 시도 오류({sel}): {e}")
                                continue
                        if not done:
                            try:
                                links = row_loc.locator("a")
                                info = [f"text={links.nth(i).inner_text()},onclick={links.nth(i).get_attribute('onclick')}"
                                        for i in range(links.count())]
                                self.log(f"  {bk.booking_no} 삭제 버튼 못 찾음. links: {info}")
                                bk.fail_reason = "삭제 버튼 못 찾음"  # v4
                            except Exception as e:
                                self.log(f"  {bk.booking_no} 디버깅 실패: {e}")
                                bk.fail_reason = f"삭제 실패: {e}"  # v4

                    elif bk.action == "modify":
                        # 재요청수량 입력 — 마지막 텍스트 input
                        try:
                            inputs = row_loc.locator("input[type='text'], input:not([type])")
                            if inputs.count():
                                inputs.last.fill(str(bk.new_qty))
                                self.log(f"  재요청수량 {bk.new_qty} 입력")
                            else:
                                # 모든 input 시도
                                row_loc.locator("input").last.fill(str(bk.new_qty))
                        except Exception as e:
                            self.log(f"  수량 입력 실패: {e}")
                        page.wait_for_timeout(300)

                        # 수정 버튼 — 확인2 컬럼
                        done = False
                        for sel in [
                            "a:has-text('수정')",
                            "a[onclick*='Re']",
                            "span.btn_pack6 a",
                            "input[value='수정']",
                            "button:has-text('수정')",
                        ]:
                            try:
                                btn = row_loc.locator(sel).last
                                if btn.count() and btn.is_visible(timeout=1000):
                                    def handle_dialog(dialog):
                                        dialog.accept()
                                    page.once("dialog", handle_dialog)
                                    btn.click()
                                    page.wait_for_timeout(2000)
                                    self.log(f"  {bk.booking_no} 수정 완료 (재요청={bk.new_qty})")
                                    bk.action = "done_modify"
                                    done = True
                                    break
                            except Exception as e:
                                self.log(f"  수정 시도 오류({sel}): {e}")
                                continue
                        if not done:
                            try:
                                inputs = row_loc.locator("input")
                                info = []
                                for i in range(inputs.count()):
                                    v = inputs.nth(i).get_attribute("value")
                                    t = inputs.nth(i).get_attribute("type")
                                    info.append(f"type={t},value={v}")
                                self.log(f"  {bk.booking_no} 수정 버튼 못 찾음. inputs: {info}")
                                bk.fail_reason = "수정 버튼 못 찾음"  # v4
                            except Exception as e:
                                self.log(f"  {bk.booking_no} 디버깅 실패: {e}")
                                bk.fail_reason = f"수정 실패: {e}"  # v4

                browser.close()

            # ══════════════════════════════════════════
            # v4: 삭제/수정 작업이 모두 끝난 마지막 단계에서만 실행되는 Hook.
            #     ⑤ Excel 보고서 생성 ⑥ Outlook 메일 작성 ⑦ Draft(임시보관함) 저장
            #     — 기존 ①~④ 처리 로직은 위에서 이미 완료된 상태이며 수정하지 않았음.
            #       실패해도 self._running/버튼 상태 복원 흐름(finally)에는 영향 없음.
            # ══════════════════════════════════════════
            try:
                send_report_mail(self.bookings, targets, log_func=self.log)
            except Exception as e:
                self.log(f"보고서/메일 처리 오류: {e}")

            self.log("=== 처리 완료 ===")
            self.after(0, self._refresh_table)

        except Exception as e:
            self.log(f"처리 오류: {e}")
        finally:
            self._running = False
            self.after(0, lambda: self.btn_etd_all.config(state="normal"))
            self.after(0, lambda: self.btn_etd.config(state="normal"))
            self.after(0, lambda: self.btn_apply.config(state="normal", text="④ 선택 항목 처리"))
            self.after(0, lambda: self.btn_apply_all.config(state="normal", text="⑤ 전체 처리"))

    # ── ⑥ 보고서 Outlook 임시보관함 저장 (수동) ────────────────
    #     v4: _process_thread 끝의 자동 Hook과는 별개로,
    #     사용자가 원하는 시점에 직접 눌러 현재 self.bookings 상태 기준으로
    #     보고서를 만들고 Outlook Draft에 저장할 수 있는 버튼.
    #     기존 ①~⑤ 로직/버튼은 전혀 건드리지 않았다.
    def _start_report_mail(self):
        if self._running:
            messagebox.showwarning("처리 중", "다른 작업이 진행 중입니다. 완료 후 다시 시도해주세요.")
            return
        if not self.bookings:
            messagebox.showinfo("없음", "먼저 ① 선반입 부킹 조회를 실행해주세요.")
            return
        self._running = True
        self.btn_mail_draft.config(state="disabled", text="Outlook 저장 중...")
        threading.Thread(target=self._report_mail_thread, daemon=True).start()

    def _report_mail_thread(self):
        try:
            targets = [b for b in self.bookings
                       if b.action in ("modify", "delete", "done_delete", "done_modify")]
            self.log(f"보고서 Outlook 임시보관함 저장 시작 (대상 {len(targets)}건)...")
            ok = send_report_mail(self.bookings, targets, log_func=self.log)
            if not ok:
                self.log("보고서 Outlook 임시보관함 저장 실패")
        except Exception as e:
            self.log(f"보고서/메일 처리 오류: {e}")
        finally:
            self._running = False
            self.after(0, lambda: self.btn_mail_draft.config(
                state="normal", text="⑥ 보고서 Outlook 임시보관함 저장"))


if __name__ == "__main__":
    try:
        app = App()
        if not BPT_ID or not BPT_PW:
            messagebox.showwarning(
                "BPT 계정 설정 필요",
                "프로그램 폴더의 config.json에 BPT_ID와 BPT_PW를 입력한 뒤 다시 실행해 주세요.\n"
                "(config.example.json을 복사해 config.json으로 저장하면 됩니다.)",
            )
        app.mainloop()
    except Exception as e:
        import traceback
        with open("app_error.txt", "w", encoding="utf-8") as f:
            f.write(traceback.format_exc())
        input(f"오류: {e}\nEnter 눌러서 닫기...")
