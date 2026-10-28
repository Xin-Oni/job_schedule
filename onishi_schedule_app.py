import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import threading
import pdfplumber
import unicodedata
import re
import os
import webbrowser
import urllib.parse
from pathlib import Path
from datetime import date, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

try:
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build as _gcal_build
    GCAL_AVAILABLE = True
except ImportError:
    GCAL_AVAILABLE = False

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# ── 定数 ──────────────────────────────────────────────────────────
TARGET_NAME = "大西"
GCAL_SCOPES = ["https://www.googleapis.com/auth/calendar.events"]
SEND_TO = "shin0428@gmail.com"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465

SHIFT_TIMED = {
    "A": ("08:00", "20:00"),
    "C": ("20:00", "32:00"),
}
# シフト表示ラベル（コード付き）
SHIFT_DISPLAY = {
    "A": "A（日勤 8:00-20:00）",
    "C": "C（夜勤 20:00-翌8:00）",
    "公": "公（公休）",
    "年": "年（年休）",
    "病": "病（病休）",
    "特": "特（特休）",
    "育": "育（育休）",
}
WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]

# ── 日本の祝日計算 ────────────────────────────────────────────────
def _nth_monday(year, month, n):
    d = date(year, month, 1)
    offset = (7 - d.weekday()) % 7
    first_mon = d + timedelta(days=offset)
    return first_mon + timedelta(weeks=n - 1)

def _equinox_day(year, base, factor):
    return int(base + factor * (year - 1980) - (year - 1980) // 4)

def jp_holidays(year: int) -> set:
    h = set()
    def add(m, d):
        try: h.add(date(year, m, d))
        except ValueError: pass

    # 固定祝日
    add(1,  1)   # 元旦
    add(2, 11)   # 建国記念の日
    add(2, 23)   # 天皇誕生日
    add(4, 29)   # 昭和の日
    add(5,  3)   # 憲法記念日
    add(5,  4)   # みどりの日
    add(5,  5)   # こどもの日
    add(8, 11)   # 山の日
    add(11, 3)   # 文化の日
    add(11,23)   # 勤労感謝の日

    # 春分・秋分（計算式）
    spring = _equinox_day(year, 20.8431, 0.242194)
    autumn = _equinox_day(year, 23.2488, 0.242194)
    add(3, spring)
    add(9, autumn)

    # ハッピーマンデー
    h.add(_nth_monday(year, 1, 2))   # 成人の日
    h.add(_nth_monday(year, 7, 3))   # 海の日
    h.add(_nth_monday(year, 9, 3))   # 敬老の日
    h.add(_nth_monday(year,10, 2))   # スポーツの日

    # 振替休日（日曜祝日 → 翌月曜、連続する場合は次の平日）
    extra = set()
    for hd in sorted(h):
        if hd.weekday() == 6:  # Sunday
            sub = hd + timedelta(days=1)
            while sub in h or sub in extra:
                sub += timedelta(days=1)
            extra.add(sub)
    h.update(extra)

    # 国民の休日（祝日に挟まれた平日）
    hlist = sorted(h)
    sandwiched = set()
    for i in range(len(hlist) - 1):
        a, b = hlist[i], hlist[i+1]
        if (b - a).days == 2:
            mid = a + timedelta(days=1)
            if mid.weekday() < 6 and mid not in h:
                sandwiched.add(mid)
    h.update(sandwiched)
    return h

_HOLIDAY_CACHE = {}
def is_jp_holiday(d: date) -> bool:
    if d.year not in _HOLIDAY_CACHE:
        _HOLIDAY_CACHE[d.year] = jp_holidays(d.year)
    return d in _HOLIDAY_CACHE[d.year]


# ── ユーティリティ ─────────────────────────────────────────────────
def _nkc(s) -> str:
    if not s:
        return ""
    return unicodedata.normalize("NFKC", str(s)).strip()

def _to_shift_key(cell_text: str) -> str:
    t = _nkc(cell_text).upper()
    if not t:
        return ""
    if t.startswith("A") or "日勤" in t or "A勤" in t:
        return "A"
    if t.startswith("C") or "夜勤" in t or "C勤" in t or "ナイト" in t:
        return "C"
    for code in ("公", "年", "病", "特", "育"):
        if t.startswith(code):
            return code
    return t[0]

def _shift_display(shift_key: str, raw: str) -> str:
    if shift_key in SHIFT_DISPLAY:
        return SHIFT_DISPLAY[shift_key]
    r = _nkc(raw)
    if r:
        return f"{r}（休み）" if shift_key == "" else r
    return "休み"

def _row_tag(d: date) -> str:
    """表示行の色タグを返す"""
    if is_jp_holiday(d):
        return "holiday"
    wd = d.weekday()
    if wd == 6:
        return "sunday"
    if wd == 5:
        return "saturday"
    return "weekday"

# ── PDFから年月を自動検出 ─────────────────────────────────────────
def detect_base_month(pdf_path: Path, tables) -> tuple:
    """(year, month) を返す。ファイル名→PDFヘッダーの順で試みる"""
    fname = pdf_path.stem

    # 1) ファイル名から YYYYMM を抽出（例: 202604TKHER...）
    m = re.search(r'(20\d{2})(\d{2})', fname)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return y, mo

    # 2) PDFのヘッダー行から「N月」または「YYYY年N月」を探す
    for table in tables:
        if not table:
            continue
        for row in table[:6]:
            if not row:
                continue
            for cell in row:
                if not cell:
                    continue
                s = str(cell)
                # "YYYY年N月" 形式
                m2 = re.search(r'(20\d{2})年\s*(\d{1,2})月', s)
                if m2:
                    return int(m2.group(1)), int(m2.group(2))
                # "N月期" 形式（年はファイル名から補完）
                m3 = re.search(r'(\d{1,2})月', s)
                if m3:
                    mo = int(m3.group(1))
                    if 1 <= mo <= 12:
                        y_m = re.search(r'(20\d{2})', fname)
                        y = int(y_m.group(1)) if y_m else date.today().year
                        return y, mo

    # 3) フォールバック：現在の年月
    return date.today().year, date.today().month

# ── PDF解析 ───────────────────────────────────────────────────────
def parse_pdf(file_path: Path, debug_log: list = None):
    def log(msg):
        if debug_log is not None:
            debug_log.append(msg)

    with pdfplumber.open(str(file_path)) as pdf:
        log(f"ページ数: {len(pdf.pages)}")
        all_tables = []
        for page in pdf.pages:
            all_tables.extend(page.extract_tables() or [])

        base_year, base_month = detect_base_month(file_path, all_tables)
        log(f"検出した基準年月: {base_year}年{base_month}月")

        raw_rows = []
        for tbl_i, table in enumerate(all_tables):
            if not table:
                continue
            n_cols = max((len(r) for r in table if r), default=0)
            log(f"\nT{tbl_i+1}: {len(table)}行×{n_cols}列")
            for ri, row in enumerate(table[:5]):
                if row:
                    log(f"  行{ri+1}: {[_nkc(c) for c in row]}")

            # 大西列を検索
            target_col = None
            header_row_idx = None
            for ri, row in enumerate(table):
                if not row:
                    continue
                for ci, cell in enumerate(row):
                    if cell and TARGET_NAME in str(cell):
                        target_col = ci
                        header_row_idx = ri
                        log(f"  ★ 大西列: 行{ri+1} 列{ci+1}")
                        break
                if target_col is not None:
                    break

            if target_col is None:
                log("  ✗ 大西列未発見")
                continue

            data_start = header_row_idx + 1

            # 日付列を検出（全データ行で数値1-31が最多の列）
            day_counts = {}
            for row in table[data_start:]:
                if not row:
                    continue
                for ci, cell in enumerate(row):
                    try:
                        v = int(_nkc(cell).split("\n")[0])
                        if 1 <= v <= 31:
                            day_counts[ci] = day_counts.get(ci, 0) + 1
                    except (ValueError, TypeError):
                        pass
            date_col = max(day_counts, key=lambda k: day_counts[k]) if day_counts else 0
            log(f"  日付列: {date_col+1}列目")

            # 全データ行を抽出（seen_days廃止 → 月またぎを正しく処理）
            log(f"  データ行サンプル（先頭5行）:")
            for ri, row in enumerate(table[data_start:data_start+5]):
                if row:
                    log(f"    データ{ri+1}: {[_nkc(c) for c in row]}")

            prev_extracted_day = -1
            for row in table[data_start:]:
                if not row:
                    continue
                clean = [_nkc(c) for c in row]
                try:
                    day = int(clean[date_col].split("\n")[0]) if date_col < len(clean) else -1
                    if not (1 <= day <= 31):
                        continue
                except (ValueError, IndexError):
                    continue
                cell_raw = clean[target_col] if target_col < len(clean) else ""
                shift_key = _to_shift_key(cell_raw)
                raw_rows.append((day, shift_key, cell_raw))
                log(f"    day={day} shift={repr(shift_key)} raw={repr(cell_raw)}")
                prev_extracted_day = day

    log(f"\n抽出合計: {len(raw_rows)}件")
    if not raw_rows:
        return [], base_year, base_month

    # 月またぎ処理
    cur_year, cur_month = base_year, base_month
    if raw_rows[0][0] > 20:
        cur_month -= 1
        if cur_month == 0:
            cur_month = 12
            cur_year -= 1

    shifts = []
    prev_day = raw_rows[0][0]
    for day, shift_key, shift_raw in raw_rows:
        if day < prev_day and prev_day > 20:
            cur_month += 1
            if cur_month == 13:
                cur_month = 1
                cur_year += 1
        prev_day = day
        try:
            work_date = date(cur_year, cur_month, day)
        except ValueError:
            continue
        shifts.append((work_date, shift_key, shift_raw))

    # ── 欠落日を「休み」で補完 ─────────────────────────────────────
    # PDFにない日付（日付列が空の休み行）を日付範囲から自動補完する
    if shifts:
        shift_date_set = {s[0] for s in shifts}
        min_date = min(shift_date_set)
        max_date = max(shift_date_set)
        fill = []
        d = min_date
        while d <= max_date:
            if d not in shift_date_set:
                fill.append((d, "", ""))   # 休み
            d += timedelta(days=1)
        if fill:
            log(f"欠落日補完: {len(fill)}件（休み扱い）")
        shifts = sorted(shifts + fill, key=lambda x: x[0])

    log(f"最終合計: {len(shifts)}件")
    return shifts, base_year, base_month


# ── Google Calendar API ───────────────────────────────────────────
def _get_gcal_service(app_dir: Path):
    """OAuth2認証してGoogle Calendar serviceを返す。(error_msg, service)"""
    creds_file = app_dir / "credentials.json"
    token_file = app_dir / "token.json"
    if not creds_file.exists():
        return None, (
            "credentials.json が見つかりません。\n\n"
            "【初回設定手順】\n"
            "1. https://console.cloud.google.com/ を開く\n"
            "2. プロジェクト作成 → Google Calendar API を有効化\n"
            "3. 認証情報 → OAuth 2.0 クライアントID を作成\n"
            "   （アプリの種類: デスクトップアプリ）\n"
            "4. JSONをダウンロードし「credentials.json」という名前で\n"
            f"   このフォルダ ({app_dir}) に保存\n\n"
            "設定後、再度ボタンを押してください。"
        )
    creds = None
    if token_file.exists():
        creds = Credentials.from_authorized_user_file(str(token_file), GCAL_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                str(creds_file), GCAL_SCOPES)
            creds = flow.run_local_server(port=0)
        token_file.write_text(creds.to_json(), encoding="utf-8")
    service = _gcal_build("calendar", "v3", credentials=creds)
    return service, None


def _make_gcal_event(work_date, shift_key, shift_raw):
    """Google Calendar API用イベントdictを生成"""
    summary = _shift_display(shift_key, shift_raw)
    if shift_key in SHIFT_TIMED:
        sh, sm = map(int, SHIFT_TIMED[shift_key][0].split(":"))
        eh_raw, em = map(int, SHIFT_TIMED[shift_key][1].split(":"))
        if eh_raw >= 24:
            eh = eh_raw - 24
            dtend = work_date + timedelta(days=1)
        else:
            eh = eh_raw
            dtend = work_date
        return {
            "summary": summary,
            "start": {"dateTime": f"{work_date.isoformat()}T{sh:02d}:{sm:02d}:00",
                      "timeZone": "Asia/Tokyo"},
            "end":   {"dateTime": f"{dtend.isoformat()}T{eh:02d}:{em:02d}:00",
                      "timeZone": "Asia/Tokyo"},
        }
    else:
        return {
            "summary": summary,
            "start": {"date": work_date.isoformat()},
            "end":   {"date": (work_date + timedelta(days=1)).isoformat()},
        }


# ── ICS生成 ───────────────────────────────────────────────────────
def make_ics(shifts) -> str:
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0",
             "PRODID:-//OhishiSchedule//JP",
             "CALSCALE:GREGORIAN", "METHOD:PUBLISH"]
    for work_date, shift_key, shift_raw in shifts:
        summary = _shift_display(shift_key, shift_raw)
        uid = f"{work_date:%Y%m%d}-{shift_key or 'X'}@onishi"
        if shift_key in SHIFT_TIMED:
            sh, sm = map(int, SHIFT_TIMED[shift_key][0].split(":"))
            eh_raw, em = map(int, SHIFT_TIMED[shift_key][1].split(":"))
            dtstart = work_date
            if eh_raw >= 24:
                eh, dtend = eh_raw - 24, work_date + timedelta(days=1)
            else:
                eh, dtend = eh_raw, work_date
            lines += ["BEGIN:VEVENT", f"UID:{uid}",
                      f"DTSTART:{dtstart:%Y%m%d}T{sh:02d}{sm:02d}00",
                      f"DTEND:{dtend:%Y%m%d}T{eh:02d}{em:02d}00",
                      f"SUMMARY:{summary}", "END:VEVENT"]
        else:
            lines += ["BEGIN:VEVENT", f"UID:{uid}",
                      f"DTSTART;VALUE=DATE:{work_date:%Y%m%d}",
                      f"DTEND;VALUE=DATE:{work_date + timedelta(days=1):%Y%m%d}",
                      f"SUMMARY:{summary}", "END:VEVENT"]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines)

# ── PDF生成 ───────────────────────────────────────────────────────
def generate_pdf(shifts, out_path: Path, year: int, month: int):
    pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
    doc = SimpleDocTemplate(str(out_path), pagesize=A4,
                            leftMargin=15*mm, rightMargin=15*mm,
                            topMargin=20*mm, bottomMargin=20*mm)
    title_style = ParagraphStyle("title", fontName="HeiseiKakuGo-W5",
                                 fontSize=13, leading=18, spaceAfter=4)
    elements = [
        Paragraph(f"大西 勤務スケジュール　{year}年{month:02d}月", title_style),
        Spacer(1, 4*mm)
    ]
    data = [["日付", "曜", "シフト"]]
    row_colors = []
    for work_date, shift_key, shift_raw in sorted(shifts, key=lambda x: x[0]):
        label = _shift_display(shift_key, shift_raw)
        data.append([work_date.strftime("%Y/%m/%d"),
                     WEEKDAY_JP[work_date.weekday()], label])
        tag = _row_tag(work_date)
        if tag == "holiday" or tag == "sunday":
            row_colors.append(colors.HexColor("#FDECEA"))
        elif tag == "saturday":
            row_colors.append(colors.HexColor("#E8F0FE"))
        else:
            row_colors.append(colors.white)

    tbl = Table(data, colWidths=[40*mm, 12*mm, 113*mm])
    style_cmds = [
        ("FONTNAME", (0,0), (-1,-1), "HeiseiKakuGo-W5"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1565C0")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.4, colors.grey),
        ("ROWHEIGHT", (0,0), (-1,-1), 7*mm),
    ]
    for i, bg in enumerate(row_colors):
        style_cmds.append(("BACKGROUND", (0, i+1), (-1, i+1), bg))
    tbl.setStyle(TableStyle(style_cmds))
    elements.append(tbl)
    doc.build(elements)

# ── HTML本文生成 ──────────────────────────────────────────────────
def build_html_body(shifts, year: int, month: int) -> str:
    rows_html = ""
    for work_date, shift_key, shift_raw in sorted(shifts, key=lambda x: x[0]):
        wd = WEEKDAY_JP[work_date.weekday()]
        label = _shift_display(shift_key, shift_raw)
        tag = _row_tag(work_date)
        if tag in ("holiday", "sunday"):
            bg = "#FDECEA"
        elif tag == "saturday":
            bg = "#E8F0FE"
        else:
            bg = "#F9F9F9"
        rows_html += (f'<tr style="background:{bg}">'
                      f'<td style="padding:5px 10px;border:1px solid #ddd">'
                      f'{work_date:%Y/%m/%d}</td>'
                      f'<td style="padding:5px 10px;border:1px solid #ddd;'
                      f'text-align:center">{wd}</td>'
                      f'<td style="padding:5px 10px;border:1px solid #ddd">'
                      f'{label}</td></tr>\n')
    return (f'<!DOCTYPE html><html><head><meta charset="utf-8">'
            f'<style>body{{font-family:sans-serif;margin:20px;color:#333}}'
            f'h2{{color:#1565C0}}table{{border-collapse:collapse;min-width:500px}}'
            f'th{{background:#1565C0;color:#fff;padding:8px 12px;'
            f'border:1px solid #1565C0}}</style></head><body>'
            f'<h2>大西 勤務スケジュール　{year}年{month:02d}月</h2>'
            f'<table><tr><th>日付</th><th>曜日</th><th>シフト</th></tr>'
            f'{rows_html}</table>'
            f'<p style="font-size:0.8em;color:#999">Generated by OnishiScheduleApp</p>'
            f'</body></html>')

# ── デバッグウィンドウ ─────────────────────────────────────────────
def show_debug_window(parent, text: str):
    win = tk.Toplevel(parent)
    win.title("PDF解析デバッグ情報")
    win.geometry("700x480")
    tk.Label(win, text="PDFの構造解析ログ",
             font=("Meiryo", 10, "bold"), pady=6).pack()
    txt = scrolledtext.ScrolledText(win, font=("Courier", 9), wrap="word")
    txt.pack(fill="both", expand=True, padx=10, pady=(0, 8))
    txt.insert("end", text)
    txt.config(state="disabled")
    tk.Button(win, text="閉じる", command=win.destroy,
              padx=20, pady=4).pack(pady=(0, 8))


# ── メインアプリ ──────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("大西 勤務スケジュール管理")
        self.geometry("700x640")
        self.resizable(True, True)
        self.configure(bg="#F0F4F8")
        self._shifts = []
        self._pdf_paths = []
        self._out_pdf_path = None
        self._sending = False
        self._last_debug = ""
        self._base_year = date.today().year
        self._base_month = date.today().month
        self.html_body = ""
        self._build_ui()

    def _build_ui(self):
        BG = "#F0F4F8"
        HDR = "#1565C0"
        BTN = "#1976D2"
        MLR = "#AD1457"
        FT  = ("Meiryo", 10)
        FB  = ("Meiryo", 10, "bold")

        # ヘッダー
        hdr = tk.Frame(self, bg=HDR); hdr.pack(fill="x")
        tk.Label(hdr, text="🏥  大西 勤務スケジュール管理",
                 font=("Meiryo", 15, "bold"), fg="white", bg=HDR,
                 pady=11).pack()

        # ── 操作エリア ──
        cfg = tk.Frame(self, bg=BG, padx=20, pady=8); cfg.pack(fill="x")

        # PDF選択
        r1 = tk.Frame(cfg, bg=BG); r1.pack(fill="x", pady=3)
        tk.Button(r1, text="📂  PDFファイルを選択", font=FB,
                  bg=BTN, fg="white", relief="flat", padx=12, pady=6,
                  command=self._select_pdf).pack(side="left")
        self.pdf_label = tk.Label(r1, text="（未選択）",
                                  font=FT, bg=BG, fg="#555")
        self.pdf_label.pack(side="left", padx=10)

        # 解析ボタン + 検出年月表示
        r2 = tk.Frame(cfg, bg=BG); r2.pack(fill="x", pady=3)
        tk.Button(r2, text="🔍  スケジュール解析", font=FB,
                  bg=BTN, fg="white", relief="flat", padx=12, pady=6,
                  command=self._parse).pack(side="left")
        self.month_label = tk.Label(r2, text="",
                                    font=("Meiryo", 10, "bold"),
                                    bg=BG, fg="#1565C0")
        self.month_label.pack(side="left", padx=14)
        tk.Button(r2, text="🔄 リセット", font=("Meiryo", 9),
                  bg="#78909C", fg="white", relief="flat", padx=6, pady=6,
                  command=self._reset).pack(side="right")

        # 凡例
        leg = tk.Frame(cfg, bg=BG); leg.pack(fill="x", pady=(2, 0))
        for bg_c, txt in [("#FDECEA","日祝"), ("#E8F0FE","土"),
                           ("#F5F5F5","休日")]:
            lf = tk.Frame(leg, bg=bg_c, relief="solid",
                          highlightbackground="#ccc", highlightthickness=1)
            lf.pack(side="left", padx=3)
            tk.Label(lf, text=f" {txt} ", font=("Meiryo", 8),
                     bg=bg_c).pack()

        # ── 結果ツリー ──
        rf = tk.Frame(self, bg=BG, padx=20); rf.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("Sched.Treeview", font=("Meiryo", 10), rowheight=22)
        style.configure("Sched.Treeview.Heading",
                        font=("Meiryo", 10, "bold"))

        self.tree = ttk.Treeview(rf, style="Sched.Treeview",
                                 columns=("date","wd","shift"),
                                 show="headings", height=16)
        self.tree.heading("date",  text="日付")
        self.tree.heading("wd",    text="曜")
        self.tree.heading("shift", text="シフト")
        self.tree.column("date",  width=115, anchor="center")
        self.tree.column("wd",    width=38,  anchor="center")
        self.tree.column("shift", width=300, anchor="w")

        self.tree.tag_configure("holiday",  background="#FDECEA")
        self.tree.tag_configure("sunday",   background="#FDECEA")
        self.tree.tag_configure("saturday", background="#E8F0FE")
        self.tree.tag_configure("yasumi",   background="#F5F5F5")
        self.tree.tag_configure("weekday",  background="white")

        sb = ttk.Scrollbar(rf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # ── ボタン行 ──
        bf = tk.Frame(self, bg=BG, padx=20, pady=8); bf.pack(fill="x")
        tk.Button(bf, text="📅  Googleカレンダーに登録",
                  font=FB, bg=BTN, fg="white", relief="flat",
                  padx=10, pady=6, command=self._upload_gcal).pack(
                  side="left", padx=(0,6))
        tk.Button(bf, text="🖨  PDF保存",
                  font=FB, bg=BTN, fg="white", relief="flat",
                  padx=10, pady=6, command=self._save_pdf).pack(
                  side="left", padx=(0,6))
        self.mail_btn = tk.Button(bf, text="📧  メール送信",
                  font=FB, bg=MLR, fg="white", relief="flat",
                  padx=10, pady=6, command=self._send_mail)
        self.mail_btn.pack(side="left")

        # ステータスバー
        self.status_var = tk.StringVar(value="PDFファイルを選択してください")
        tk.Label(self, textvariable=self.status_var,
                 font=("Meiryo", 9), bg="#DDE3EA",
                 anchor="w", padx=10, pady=3).pack(fill="x", side="bottom")


    # ── アクション ────────────────────────────────────────────────
    def _select_pdf(self):
        paths = filedialog.askopenfilenames(
            title="PDFファイルを選択",
            filetypes=[("PDF", "*.pdf"), ("All", "*.*")])
        if paths:
            self._pdf_paths = [Path(p) for p in paths]
            names = ", ".join(p.name for p in self._pdf_paths)
            self.pdf_label.config(
                text=names[:65] + ("…" if len(names) > 65 else ""))
            self.status_var.set(f"{len(self._pdf_paths)}件のPDFを選択")
            self.after(100, self._parse)  # 選択後すぐ自動解析

    def _parse(self):
        if not self._pdf_paths:
            messagebox.showwarning("警告", "PDFファイルを先に選択してください")
            return
        self.status_var.set("解析中…"); self.update()

        all_shifts = []
        all_debug = []
        last_year, last_month = self._base_year, self._base_month

        for p in self._pdf_paths:
            dbg = [f"=== {p.name} ==="]
            try:
                shifts, yr, mo = parse_pdf(p, dbg)
                all_shifts.extend(shifts)
                last_year, last_month = yr, mo
                dbg.append(f"→ {len(shifts)}件抽出")
            except Exception as e:
                dbg.append(f"ERROR: {e}")
                messagebox.showerror("解析エラー", f"{p.name}:\n{e}")
                self.status_var.set("解析失敗")
                self._last_debug = "\n".join(dbg)
                return
            all_debug.extend(dbg)
        self._last_debug = "\n".join(all_debug)
        self._base_year, self._base_month = last_year, last_month

        # 重複排除（同日は最初のエントリを優先）
        seen = set()
        unique = []
        for s in sorted(all_shifts, key=lambda x: x[0]):
            if s[0] not in seen:
                seen.add(s[0])
                unique.append(s)
        self._shifts = unique
        self._out_pdf_path = None

        # 検出年月を表示
        self.month_label.config(
            text=f"📅 {last_year}年{last_month:02d}月 ({len(unique)}件)")

        # ツリーを更新
        for item in self.tree.get_children():
            self.tree.delete(item)

        for work_date, shift_key, shift_raw in self._shifts:
            wd_str = WEEKDAY_JP[work_date.weekday()]
            label = _shift_display(shift_key, shift_raw)
            base_tag = _row_tag(work_date)

            # シフト固有色と曜日・祝日色の優先順位
            if base_tag in ("holiday", "sunday"):
                tag = "holiday"
            elif base_tag == "saturday":
                tag = "saturday"
            elif shift_key in ("公","年","病","特","育"):
                tag = "yasumi"
            else:
                tag = "weekday"

            self.tree.insert("", "end",
                values=(work_date.strftime("%Y/%m/%d"), wd_str, label),
                tags=(tag,))

        self.html_body = build_html_body(
            self._shifts, last_year, last_month)

        n_a = sum(1 for s in self._shifts if s[1] == "A")
        n_c = sum(1 for s in self._shifts if s[1] == "C")
        self.status_var.set(
            f"解析完了：{len(unique)}件　日勤 {n_a}件　夜勤 {n_c}件　"
            f"その他 {len(unique)-n_a-n_c}件")

        if len(unique) == 0:
            messagebox.showwarning("0件", "シフトが検出されませんでした。\nPDFファイルを確認してください。")

    def _reset(self):
        self._pdf_paths = []
        self._shifts = []
        self._out_pdf_path = None
        self._last_debug = ""
        self._base_year = date.today().year
        self._base_month = date.today().month
        self.html_body = ""
        self.pdf_label.config(text="PDFファイルが選択されていません")
        self.month_label.config(text="")
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status_var.set("PDFファイルを選択してください")

    def _save_ics(self):
        if not self._shifts:
            messagebox.showwarning("警告", "先にスケジュールを解析してください")
            return
        y, m = self._base_year, self._base_month
        out = filedialog.asksaveasfilename(
            title="ICSファイルを保存", defaultextension=".ics",
            initialfile=f"schedule_{y}{m:02d}.ics",
            filetypes=[("iCalendar", "*.ics"), ("All", "*.*")])
        if not out:
            return
        Path(out).write_text(make_ics(self._shifts), encoding="utf-8")
        self.status_var.set(f"ICS保存: {Path(out).name}")
        if messagebox.askyesno("確認", "ICSを保存しました。\nGoogleカレンダーを開きますか？"):
            webbrowser.open(
                "https://calendar.google.com/calendar/r/settings/export")

    def _upload_gcal(self):
        if not self._shifts:
            messagebox.showwarning("警告", "先にスケジュールを解析してください")
            return
        y, m = self._base_year, self._base_month

        # ── ICSをDownloadsに自動保存（ダイアログなし）──
        dl = Path.home() / "Downloads"
        if not dl.exists():
            dl = Path.home() / "ダウンロード"
        if not dl.exists():
            dl = Path(__file__).parent
        ics_name = f"schedule_{y}{m:02d}.ics"
        ics_path = dl / ics_name
        try:
            ics_path.write_text(make_ics(self._shifts), encoding="utf-8")
        except Exception as e:
            messagebox.showerror("保存エラー", f"ICSファイルの保存に失敗しました:\n{e}")
            return

        # ── クリップボードにパスをコピー ──
        try:
            self.clipboard_clear()
            self.clipboard_append(str(ics_path))
        except Exception:
            pass

        # ── Downloadsフォルダを開く ──
        try:
            os.startfile(str(dl))
        except Exception:
            pass

        # ── Googleカレンダーのインポート画面を開く ──
        webbrowser.open("https://calendar.google.com/calendar/r/settings/export?authuser=shin0428%40gmail.com")

        self.status_var.set(f"ICS保存済み: {ics_path}  ブラウザで「ファイルを選択」→「インポート」")
        messagebox.showinfo(
            "あと2クリックで完了",
            f"保存しました:\n  {ics_path}\n\n"
            "ブラウザに切り替えて:\n"
            "① 「ファイルを選択」をクリック\n"
            "② 上のファイルを選択\n"
            "③ 「インポート」をクリック\n\n"
            "※ファイルパスはクリップボードにコピー済みです"
        )

    def _save_pdf(self):
        if not self._shifts:
            messagebox.showwarning("警告", "先にスケジュールを解析してください")
            return
        y, m = self._base_year, self._base_month
        out = filedialog.asksaveasfilename(
            title="PDFを保存", defaultextension=".pdf",
            initialfile=f"onishi_{y}{m:02d}.pdf",
            filetypes=[("PDF", "*.pdf"), ("All", "*.*")])
        if not out:
            return
        try:
            generate_pdf(self._shifts, Path(out), y, m)
            self._out_pdf_path = Path(out)
            self.status_var.set(f"PDF保存: {Path(out).name}")
        except Exception as e:
            messagebox.showerror("PDF生成エラー", str(e))

    def _send_mail(self):
        if self._sending:
            return
        if not self._shifts:
            messagebox.showwarning("警告", "先にスケジュールを解析してください")
            return
        if self._out_pdf_path is None or not self._out_pdf_path.exists():
            y, m = self._base_year, self._base_month
            import tempfile
            self._out_pdf_path = (Path(tempfile.gettempdir())
                                  / f"onishi_{y}{m:02d}.pdf")
            try:
                generate_pdf(self._shifts, self._out_pdf_path, y, m)
            except Exception as e:
                messagebox.showerror("PDF生成エラー", str(e)); return

        self._sending = True
        self.mail_btn.config(state="disabled", text="送信中…", bg="#888")
        self.status_var.set("メール送信中…")
        threading.Thread(target=self._send_worker, daemon=True).start()

    def _send_worker(self):
        try:
            y, m = self._base_year, self._base_month
            subject = f"大西 勤務スケジュール {y}年{m:02d}月"
            sent = self._try_outlook(subject)
            if not sent:
                sent = self._try_smtp(subject)
            if not sent:
                self._browser_fallback(subject)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("送信エラー", str(e)))
        finally:
            self.after(0, self._mail_done)

    def _mail_done(self):
        self._sending = False
        self.mail_btn.config(state="normal",
                             text="📧  メール送信", bg="#AD1457")
        self.status_var.set("メール送信完了（または手動送信）")

    def _try_outlook(self, subject):
        try:
            import win32com.client
            try:
                ol = win32com.client.GetActiveObject("Outlook.Application")
            except Exception:
                ol = win32com.client.Dispatch("Outlook.Application")
            if ol.Session.Accounts.Count == 0:
                return False
            mail = ol.CreateItem(0)
            mail.To = SEND_TO
            mail.Subject = subject
            mail.HTMLBody = self.html_body
            mail.Attachments.Add(str(self._out_pdf_path.resolve()))
            mail.SendUsingAccount = ol.Session.Accounts.Item(1)
            mail.Send()
            return True
        except Exception:
            return False

    def _try_smtp(self, subject):
        import tkinter.simpledialog as sd
        pw = sd.askstring("Gmailアプリパスワード",
                          "アプリパスワード入力（キャンセル→ブラウザ送信）",
                          show="*")
        if not pw:
            return False
        try:
            msg = MIMEMultipart()
            msg["From"] = msg["To"] = SEND_TO
            msg["Subject"] = subject
            msg.attach(MIMEText(self.html_body, "html", "utf-8"))
            part = MIMEBase("application", "octet-stream")
            part.set_payload(self._out_pdf_path.read_bytes())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f'attachment; filename="{self._out_pdf_path.name}"')
            msg.attach(part)
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as s:
                s.login(SEND_TO, pw)
                s.send_message(msg)
            return True
        except Exception:
            return False

    def _browser_fallback(self, subject):
        try:
            os.startfile(str(self._out_pdf_path))
        except Exception:
            pass
        params = urllib.parse.urlencode(
            {"to": SEND_TO, "su": subject,
             "body": "勤務スケジュールを添付しています。"})
        webbrowser.open(
            f"https://mail.google.com/mail/?view=cm&fs=1&{params}")


if __name__ == "__main__":
    App().mainloop()
