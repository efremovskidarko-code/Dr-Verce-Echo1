# -*- coding: utf-8 -*-
"""
ВЕТЕРИНАРНА АМБУЛАНТА ДР-ВЕРЧЕ  (верзија 2)
Потребни пакети (еднаш):  pip install reportlab tkcalendar
Стартување:               python ambulanta.py
Правење .exe:             pip install pyinstaller
                          pyinstaller --onefile --windowed ambulanta.py
"""
import os
import io
import re
import sys
import json
import base64
import sqlite3
import subprocess
from datetime import datetime, date, timedelta
from xml.sax.saxutils import escape

import tkinter as tk
from tkinter import ttk, messagebox

try:
    from tkcalendar import DateEntry
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                    PageBreak, Flowable)
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfbase.pdfmetrics import stringWidth
except ImportError:
    _r = tk.Tk()
    _r.withdraw()
    messagebox.showerror("Недостасуваат пакети",
                         "Отвори Command Prompt и напиши:\n\npip install reportlab tkcalendar")
    sys.exit(1)

# ----------------------------------------------------------------------------
# Основни поставки
# ----------------------------------------------------------------------------
CLINIC = "ВЕТЕРИНАРНА АМБУЛАНТА ДР-ВЕРЧЕ"
CLINIC_PDF = "ВЕТЕРИНАРНА АМБУЛАНТА ДР. ВЕРЧЕ"
CLINIC_FOOTER = "Ветеринарна амбуланта Др. Верче"

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(APP_DIR, "ambulanta.db")
BLOOD_DIR = os.path.join(APP_DIR, "Крвни слики")
REPORT_DIR = os.path.join(APP_DIR, "Извештаи")
os.makedirs(BLOOD_DIR, exist_ok=True)
os.makedirs(REPORT_DIR, exist_ok=True)

PRIMARY = "#0F766E"
DARK = "#134E4A"
ACCENT = "#14B8A6"
BG = "#F1F5F9"
TEXT = "#0F172A"
MUTED = "#64748B"
RED = "#DC2626"
BLUE = "#2563EB"
FONT = "Segoe UI"

DAYS = ["Понеделник", "Вторник", "Среда", "Четврток", "Петок", "Сабота", "Недела"]
DAYS_SHORT = ["Пон", "Вто", "Сре", "Чет", "Пет", "Саб", "Нед"]
SPECIES = ["Куче", "Маче"]
SEXES = ["", "Машко", "Женско"]

UP, DOWN = "▲", "▼"

# Логото на амбулантата (вградено)
LOGO_B64 = "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"

# ----------------------------------------------------------------------------
# Лабораториски параметри и референтни вредности
# (име, мерна единица, референца куче, референца маче)  -  None = нема референца
# Слободно менувај ги вредностите овде ако лабораторијата ги смени.
# ----------------------------------------------------------------------------
HEMATOLOGY = [
    ("Бели крвни клетки (WBC)", "10^9/l", "6,0-17,0", "5,5-19,5"),
    ("Лимфоцити", "10^9/l", "0,8-5,1", "0,8-7,0"),
    ("Моноцити", "10^9/l", "0,0-1,8", "0,0-1,9"),
    ("Гранулоцити", "10^9/l", "4,0-12,6", "2,1-15,0"),
    ("Лимфоцити %", "%", "12,0-30,0", "12,0-45,0"),
    ("Моноцити %", "%", "2,0-9,0", "2,0-9,0"),
    ("Гранулоцити %", "%", "60,0-83,0", "35,0-85,0"),
    ("Црвени крвни клетки (RBC)", "10^12/l", "5,50-8,50", "4,60-10,00"),
    ("Хемоглобин", "g/l", "110-190", "93-153"),
    ("Хематокрит", "%", "39,0-56,0", "28,0-49,0"),
    ("MCV", "fl", "62,0-72,0", "39,0-52,0"),
    ("MCH", "pg", "20,0-25,0", "13,0-21,0"),
    ("MCHC", "g/l", "300-380", "300-380"),
    ("RDW", "%", "11,0-15,5", "14,0-18,0"),
    ("MPV", "fl", "7,0-12,9", "5,0-11,8"),
    ("Крвни плочки (PLT)", "10^9/l", "117-460", "100-514"),
    ("EOS", "%", "2-10", "2-12"),
    ("PDW", "", None, None),
    ("PCT (Procalcitonin)", "%", None, None),
]

BIOCHEMISTRY = [
    ("AST", "U/L", "8.9-48.5", "9.2-40"),
    ("ALT", "U/L", "8.2-57.3", "8.3-53"),
    ("Креатинин", "µmol/l", "44.30-138.40", "49-165"),
    ("Уреа UV", "mmol/l", "3.10-9.20", "5.5-11.1"),
    ("Гликоза", "mmol/l", "3.40-6,00", "3.40-6.9"),
    ("GGT (gamma-glutamyltransferase)", "U/L", "1-12", "1-12"),
    ("Alkaline Phosphatase (ALP)", "U/L", "5-131", "12-65"),
    ("Вкупно протеини", "g/l", "60-79", "60-94"),
    ("Албумини", "g/l", "28-39", "24-56"),
    ("Глобулини", "g/l", "26-51", "24-47"),
    ("Триглицериди", "mmol/l", "0.23-1.47", "0.6-1.1"),
    ("Холестерол", "mmol/l", "3.3-9.0", "2.0-7.9"),
]

GROUPS = [("hem", "Хематолошка анализа", "ХЕМАТОЛОШКА АНАЛИЗА", HEMATOLOGY),
          ("bio", "Биохемиска анализа", "БИОХЕМИСКА АНАЛИЗА", BIOCHEMISTRY)]


def get_ref(param, species):
    """Враќа (текст за приказ, мин, макс) или None."""
    s = param[2] if species == "Куче" else param[3]
    if not s:
        return None
    try:
        lo, hi = [float(x) for x in s.replace(",", ".").split("-")]
    except ValueError:
        return None
    return s.replace("-", "–"), lo, hi


def compute_flag(v, ref):
    if v is None or ref is None:
        return ""
    if v < ref[1]:
        return DOWN
    if v > ref[2]:
        return UP
    return ""


def unit_ui(u):
    return u.replace("^12", "¹²").replace("^9", "⁹")


def unit_pdf(u):
    return re.sub(r"\^(\d+)", r"<super>\1</super>", escape(u))


# ----------------------------------------------------------------------------
# Помошни функции
# ----------------------------------------------------------------------------
def to_disp(iso):
    try:
        return datetime.strptime(iso, "%Y-%m-%d").strftime("%d.%m.%Y")
    except Exception:
        return iso or ""


def parse_num(s):
    s = (s or "").strip().replace(",", ".")
    if not s:
        return None
    return float(s)


def safe_name(s):
    s = re.sub(r'[\\/:*?"<>|]+', "", s).strip()
    return re.sub(r"\s+", "_", s)


def unique_path(folder, filename):
    base, ext = os.path.splitext(filename)
    path = os.path.join(folder, filename)
    n = 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{base}_{n}{ext}")
        n += 1
    return path


def open_file(path):
    if not path or not os.path.exists(path):
        messagebox.showerror("Грешка", "PDF датотеката не е пронајдена:\n" + str(path))
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Грешка", f"Не може да се отвори PDF:\n{e}")


def make_date(parent, value=None):
    kw = dict(date_pattern="dd.mm.yyyy", width=12, firstweekday="monday",
              background=PRIMARY, foreground="white", borderwidth=2)
    try:
        w = DateEntry(parent, locale="mk_MK", **kw)
    except Exception:
        w = DateEntry(parent, **kw)
    if value:
        w.set_date(value)
    return w


def get_date_iso(w):
    return w.get_date().strftime("%Y-%m-%d")


def time_values():
    return [f"{h:02d}:{m:02d}" for h in range(7, 22) for m in (0, 15, 30, 45)]


# ----------------------------------------------------------------------------
# PDF - изглед според моделот на амбулантата
# ----------------------------------------------------------------------------
def register_fonts():
    candidates = [
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
        ("C:/Windows/Fonts/segoeui.ttf", "C:/Windows/Fonts/segoeuib.ttf"),
        ("C:/Windows/Fonts/calibri.ttf", "C:/Windows/Fonts/calibrib.ttf"),
        ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
         "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ]
    for reg, bold in candidates:
        if os.path.exists(reg) and os.path.exists(bold):
            pdfmetrics.registerFont(TTFont("F", reg))
            pdfmetrics.registerFont(TTFont("FB", bold))
            return True
    return False


FONTS_OK = register_fonts()

NAVY = colors.HexColor("#15365B")
LOGO_BLUE = colors.HexColor("#4BA3E8")
SOFT = colors.HexColor("#F4F7FB")
LINE = colors.HexColor("#E2E8F0")
GRAY = colors.HexColor("#64748B")
INK = colors.HexColor("#1E293B")
PILL_UP_BG, PILL_UP_FG = colors.HexColor("#FEE2E2"), colors.HexColor("#DC2626")
PILL_DN_BG, PILL_DN_FG = colors.HexColor("#DBEAFE"), colors.HexColor("#2563EB")

PW, PH = A4
MX = 1.4 * cm
CW = PW - 2 * MX
LOGO_RATIO = 1192 / 1320

_imgs = {}


def logo_images():
    if "logo" not in _imgs:
        try:
            from PIL import Image
            im = Image.open(io.BytesIO(base64.b64decode(LOGO_B64))).convert("RGBA")
            wm = im.copy()
            wm.putalpha(im.split()[3].point(lambda v: int(v * 0.07)))
            _imgs["logo"], _imgs["wm"] = ImageReader(im), ImageReader(wm)
        except Exception:
            _imgs["logo"], _imgs["wm"] = None, None
    return _imgs["logo"], _imgs["wm"]


def page_decor(title, subtitle):
    def _page(c, doc):
        logo, wm = logo_images()
        c.saveState()
        if wm:
            ww = 13.4 * cm
            wh = ww * LOGO_RATIO
            c.drawImage(wm, (PW - ww) / 2, (PH - wh) / 2 - 1 * cm, ww, wh, mask="auto")
        if logo:
            lw = 3.1 * cm
            lh = lw * LOGO_RATIO
            c.drawImage(logo, MX, PH - 0.5 * cm - lh, lw, lh, mask="auto")
        tx = MX + 3.6 * cm
        c.setFillColor(NAVY)
        c.setFont("FB", 9.5)
        c.drawString(tx, PH - 1.3 * cm, CLINIC_PDF)
        size = 20
        while size > 12 and stringWidth(title, "FB", size) > PW - MX - tx:
            size -= 0.5
        c.setFont("FB", size)
        c.drawString(tx, PH - 2.4 * cm, title)
        c.setFillColor(GRAY)
        c.setFont("F", 10)
        c.drawString(tx, PH - 3.1 * cm, subtitle)
        ly = PH - 3.8 * cm
        c.setFillColor(NAVY)
        c.rect(MX, ly, 4.5 * cm, 0.1 * cm, fill=1, stroke=0)
        c.setFillColor(LOGO_BLUE)
        c.rect(MX + 4.5 * cm, ly, CW - 4.5 * cm, 0.1 * cm, fill=1, stroke=0)
        c.setFillColor(GRAY)
        c.setFont("F", 8.5)
        c.drawCentredString(PW / 2, 1.15 * cm, CLINIC_FOOTER)
        c.setFillColor(LOGO_BLUE)
        c.rect(0, 0, PW, 0.38 * cm, fill=1, stroke=0)
        c.restoreState()
    return _page


class Cards(Flowable):
    """8 картички (4 x 2) со податоци за пациентот."""
    CH = 1.35 * cm
    GAP = 0.3 * cm

    def __init__(self, items):
        super().__init__()
        self.items = items
        self.h = 2 * self.CH + self.GAP

    def wrap(self, aw, ah):
        return CW, self.h

    def draw(self):
        c = self.canv
        cw = (CW - 3 * self.GAP) / 4
        for i, (label, val) in enumerate(self.items):
            r, col = divmod(i, 4)
            x = col * (cw + self.GAP)
            y = self.h - (r + 1) * self.CH - r * self.GAP
            c.setFillColor(SOFT)
            c.setStrokeColor(LINE)
            c.roundRect(x, y, cw, self.CH, 0.25 * cm, fill=1, stroke=1)
            c.setFillColor(GRAY)
            c.setFont("F", 7.5)
            c.drawString(x + 0.35 * cm, y + self.CH - 0.5 * cm, label)
            size = 11
            while size > 7 and stringWidth(val, "FB", size) > cw - 0.7 * cm:
                size -= 0.5
            c.setFillColor(INK)
            c.setFont("FB", size)
            c.drawString(x + 0.35 * cm, y + 0.3 * cm, val)


class Bar(Flowable):
    """Тамносина лента со наслов на дел."""

    def __init__(self, text):
        super().__init__()
        self.text = text
        self.keepWithNext = 1

    def wrap(self, aw, ah):
        return CW, 0.85 * cm

    def draw(self):
        c = self.canv
        c.setFillColor(NAVY)
        c.roundRect(0, 0, CW, 0.85 * cm, 0.25 * cm, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("FB", 11)
        c.drawString(0.5 * cm, 0.28 * cm, self.text)


class Pill(Flowable):
    """Обоена 'таблетка' за вредност надвор од референцата."""

    def __init__(self, flag, value):
        super().__init__()
        self.flag, self.text = flag, f"{flag} {value}"
        self.w = max(2.2 * cm, stringWidth(self.text, "FB", 10) + 0.7 * cm)
        self.h = 0.58 * cm

    def wrap(self, aw, ah):
        return self.w, self.h

    def draw(self):
        c = self.canv
        up = self.flag == UP
        c.setFillColor(PILL_UP_BG if up else PILL_DN_BG)
        c.roundRect(0, 0, self.w, self.h, 0.16 * cm, fill=1, stroke=0)
        c.setFillColor(PILL_UP_FG if up else PILL_DN_FG)
        c.setFont("FB", 10)
        c.drawString(0.3 * cm, 0.17 * cm, self.text)


class Legend(Flowable):
    def wrap(self, aw, ah):
        return CW, 0.6 * cm

    def draw(self):
        c = self.canv
        x = 0
        for flag, txt, bg, fg in ((UP, "Над референтна вредност", PILL_UP_BG, PILL_UP_FG),
                                  (DOWN, "Под референтна вредност", PILL_DN_BG, PILL_DN_FG)):
            c.setFillColor(bg)
            c.roundRect(x, 0.02 * cm, 0.6 * cm, 0.5 * cm, 0.12 * cm, fill=1, stroke=0)
            c.setFillColor(fg)
            c.setFont("FB", 8)
            c.drawCentredString(x + 0.3 * cm, 0.2 * cm, flag)
            c.setFillColor(GRAY)
            c.setFont("F", 9.5)
            c.drawString(x + 0.85 * cm, 0.2 * cm, txt)
            x += 0.85 * cm + stringWidth(txt, "F", 9.5) + 0.9 * cm


class Signature(Flowable):
    """Потпис долу десно на последната страница."""

    def wrap(self, aw, ah):
        return 0, 0.1

    def draw(self):
        c = self.canv
        ax, ay = c.absolutePosition(0, 0)
        right = PW - MX
        y = 3.3 * cm
        c.saveState()
        c.translate(-ax, -ay)
        c.setStrokeColor(LINE)
        c.setLineWidth(0.8)
        c.line(right - 6.2 * cm, y + 0.7 * cm, right, y + 0.7 * cm)
        c.setFillColor(INK)
        c.setFont("FB", 10.5)
        c.drawRightString(right, y, "Доктор по ветеринарна медицина")
        c.setFillColor(GRAY)
        c.setFont("F", 9.5)
        c.drawRightString(right, y - 0.5 * cm, "Потпис и печат")
        c.restoreState()


def _styles():
    return {
        "th": ParagraphStyle("th", fontName="FB", fontSize=7.5, textColor=GRAY, leading=10),
        "cell": ParagraphStyle("cell", fontName="F", fontSize=10.5, textColor=INK, leading=13),
        "res": ParagraphStyle("res", fontName="FB", fontSize=10.5, textColor=INK, leading=13),
        "unit": ParagraphStyle("unit", fontName="F", fontSize=10, textColor=GRAY, leading=13),
        "p": ParagraphStyle("p", fontName="F", fontSize=10.5, textColor=INK, leading=15, spaceBefore=6,
                            spaceAfter=10),
    }


def patient_cards(p, date_iso, last_label="ДАТУМ НА ПРИМЕРОК"):
    def v(x):
        return (x or "").strip() or "—"
    return [("СОПСТВЕНИК", f"{p['owner_name']} {p['owner_surname']}"), ("ВИД", p["species"]),
            ("РАСА", v(p.get("breed"))), ("ИМЕ", p["pet_name"]),
            ("СТАРОСТ", v(p.get("age"))), ("ПОЛ", v(p.get("sex"))),
            ("МИКРОЧИП", v(p.get("microchip"))), (last_label, to_disp(date_iso))]


def result_table(rows):
    st = _styles()
    data = [[Paragraph(t, st["th"]) for t in ("ПАРАМЕТАР", "НАОД", "МЕРНА ЕДИНИЦА", "РЕФЕРЕНЦА")]]
    for r in rows:
        res = Pill(r["flag"], r["value"]) if r["flag"] else Paragraph(escape(r["value"]), st["res"])
        data.append([Paragraph(escape(r["name"]), st["cell"]), res,
                     Paragraph(unit_pdf(r["unit"]), st["unit"]), Paragraph(escape(r["ref"]), st["unit"])])
    t = Table(data, colWidths=[7.2 * cm, 3.4 * cm, 3.8 * cm, CW - 14.4 * cm])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, 0), 0.5, LINE),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [SOFT, colors.white]),
    ]))
    return t


def make_lab_pdf(patient, date_iso, hem, bio, note, path):
    title = "Резултати од лабораториска анализа"
    if hem and bio:
        sub = "Хематолошка и биохемиска анализа"
    else:
        sub = "Хематолошка анализа" if hem else "Биохемиска анализа"
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=MX, rightMargin=MX, topMargin=4.4 * cm,
                            bottomMargin=2.4 * cm, title=title)
    story = [Cards(patient_cards(patient, date_iso)), Spacer(1, 0.5 * cm)]
    for key, rows, bar in (("hem", hem, "ХЕМАТОЛОШКА АНАЛИЗА"), ("bio", bio, "БИОХЕМИСКА АНАЛИЗА")):
        if not rows:
            continue
        if key == "bio" and hem:
            story.append(PageBreak())
        story += [Bar(bar), Spacer(1, 0.15 * cm), result_table(rows)]
    story.append(Spacer(1, 0.4 * cm))
    story.append(Legend())
    if note.strip():
        story.append(Spacer(1, 0.3 * cm))
        story.append(Paragraph(escape(note).replace("\n", "<br/>"), _styles()["p"]))
    story.append(Signature())
    deco = page_decor(title, sub)
    doc.build(story, onFirstPage=deco, onLaterPages=deco)


def make_report_pdf(patient, date_iso, sections, path):
    st = _styles()
    title = "Извештај од преглед"
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=MX, rightMargin=MX, topMargin=4.4 * cm,
                            bottomMargin=2.4 * cm, title=title)
    story = [Cards(patient_cards(patient, date_iso, "ДАТУМ")), Spacer(1, 0.4 * cm)]
    for name, body in sections:
        if body.strip():
            story.append(Bar(name.upper()))
            story.append(Paragraph(escape(body).replace("\n", "<br/>"), st["p"]))
    story.append(Signature())
    deco = page_decor(title, "Наод, дијагноза и терапија")
    doc.build(story, onFirstPage=deco, onLaterPages=deco)


# ----------------------------------------------------------------------------
# База
# ----------------------------------------------------------------------------
SCHEMA = """
CREATE TABLE IF NOT EXISTS patients(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    owner_name TEXT NOT NULL,
    owner_surname TEXT NOT NULL,
    pet_name TEXT NOT NULL,
    species TEXT NOT NULL,
    phone TEXT DEFAULT '',
    breed TEXT DEFAULT '',
    age TEXT DEFAULT '',
    sex TEXT DEFAULT '',
    microchip TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS therapies(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_id INTEGER NOT NULL REFERENCES patients(id) ON DELETE CASCADE,
    date TEXT NOT NULL,
    note TEXT DEFAULT '',
    next_visit TEXT DEFAULT '',
    next_time TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS bloodtests(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_id INTEGER NOT NULL REFERENCES patients(id) ON DELETE CASCADE,
    date TEXT NOT NULL,
    data TEXT DEFAULT '',
    path TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS reports(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_id INTEGER NOT NULL REFERENCES patients(id) ON DELETE CASCADE,
    date TEXT NOT NULL,
    data TEXT DEFAULT '',
    path TEXT NOT NULL
);
"""


def migrate(con):
    cols = {r[1] for r in con.execute("PRAGMA table_info(patients)")}
    for c in ("breed", "age", "sex", "microchip"):
        if c not in cols:
            con.execute(f"ALTER TABLE patients ADD COLUMN {c} TEXT DEFAULT ''")
    con.commit()


# ----------------------------------------------------------------------------
# Дијалози
# ----------------------------------------------------------------------------
class Dialog(tk.Toplevel):
    def __init__(self, parent, title):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=BG)
        self.transient(parent)
        self.resizable(False, False)

    def show(self):
        self.update_idletasks()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        x = self.master.winfo_rootx() + (self.master.winfo_width() - w) // 2
        y = self.master.winfo_rooty() + max(10, (self.master.winfo_height() - h) // 2)
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        try:
            self.grab_set()
        except tk.TclError:
            pass

    def buttons(self, save_text="Зачувај", cmd=None):
        bar = ttk.Frame(self, padding=12)
        bar.pack(fill="x")
        ttk.Button(bar, text=save_text, style="Accent.TButton", command=cmd).pack(side="right")
        ttk.Button(bar, text="Откажи", command=self.destroy).pack(side="right", padx=8)


class PatientDialog(Dialog):
    FIELDS = [("owner_name", "Име на сопственик"), ("owner_surname", "Презиме на сопственик"),
              ("pet_name", "Име на милениче"), ("species", "Вид на животно"),
              ("breed", "Раса (по желба)"), ("age", "Старост (по желба)"),
              ("sex", "Пол (по желба)"), ("microchip", "Микрочип (по желба)"),
              ("phone", "Телефон (по желба)")]

    def __init__(self, app, patient=None):
        super().__init__(app.root, "Пациент" if patient else "Нов пациент")
        self.app, self.patient = app, patient
        f = ttk.Frame(self, padding=16)
        f.pack()
        self.vars = {}
        for i, (key, label) in enumerate(self.FIELDS):
            ttk.Label(f, text=label).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 12))
            v = tk.StringVar(value=(patient[key] or "") if patient else "")
            self.vars[key] = v
            if key == "species":
                if not patient:
                    v.set(SPECIES[0])
                ttk.Combobox(f, textvariable=v, values=SPECIES, state="readonly", width=28).grid(row=i, column=1)
            elif key == "sex":
                ttk.Combobox(f, textvariable=v, values=SEXES, state="readonly", width=28).grid(row=i, column=1)
            else:
                ttk.Entry(f, textvariable=v, width=31).grid(row=i, column=1)
        self.buttons(cmd=self.save)
        self.show()

    def save(self):
        d = {k: v.get().strip() for k, v in self.vars.items()}
        if not (d["owner_name"] and d["owner_surname"] and d["pet_name"]):
            messagebox.showwarning("Внимание", "Внеси име, презиме и име на милениче.", parent=self)
            return
        con = self.app.con
        keys = [k for k, _ in self.FIELDS]
        if self.patient:
            con.execute(f"UPDATE patients SET {','.join(k + '=?' for k in keys)} WHERE id=?",
                        [d[k] for k in keys] + [self.patient["id"]])
            pid = self.patient["id"]
        else:
            cur = con.execute(f"INSERT INTO patients({','.join(keys)}) VALUES({','.join('?' * len(keys))})",
                              [d[k] for k in keys])
            pid = cur.lastrowid
        con.commit()
        self.destroy()
        self.app.select_patient(pid)


class TherapyDialog(Dialog):
    def __init__(self, app, patient):
        super().__init__(app.root, "Нова терапија / забелешка")
        self.app, self.patient = app, patient
        f = ttk.Frame(self, padding=16)
        f.pack()
        ttk.Label(f, text="Датум").grid(row=0, column=0, sticky="w", pady=6)
        self.date = make_date(f)
        self.date.grid(row=0, column=1, sticky="w")
        ttk.Label(f, text="Терапија / забелешка").grid(row=1, column=0, sticky="nw", pady=6)
        self.text = tk.Text(f, width=52, height=8, font=(FONT, 11), wrap="word", relief="solid", bd=1)
        self.text.grid(row=1, column=1, pady=6)

        self.use_next = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Закажи следна посета", variable=self.use_next,
                        command=self.toggle).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 2))
        nf = ttk.Frame(f)
        nf.grid(row=3, column=0, columnspan=2, sticky="w")
        ttk.Label(nf, text="Следна посета").pack(side="left", padx=(0, 8))
        self.next_date = make_date(nf, date.today() + timedelta(days=1))
        self.next_date.pack(side="left")
        ttk.Label(nf, text="Час").pack(side="left", padx=(12, 6))
        self.time = ttk.Combobox(nf, values=time_values(), width=7)
        self.time.set("09:00")
        self.time.pack(side="left")
        self.nf_widgets = [self.next_date, self.time]
        self.toggle()
        self.buttons(cmd=self.save)
        self.show()

    def toggle(self):
        state = "normal" if self.use_next.get() else "disabled"
        for w in self.nf_widgets:
            w.configure(state=state)

    def save(self):
        note = self.text.get("1.0", "end").strip()
        if not note:
            messagebox.showwarning("Внимание", "Напиши забелешка.", parent=self)
            return
        nv, nt = "", ""
        if self.use_next.get():
            nv = get_date_iso(self.next_date)
            nt = self.time.get().strip()
            if not re.fullmatch(r"\d{1,2}:\d{2}", nt):
                messagebox.showwarning("Внимание", "Часот мора да е во формат 09:30", parent=self)
                return
            nt = f"{int(nt.split(':')[0]):02d}:{nt.split(':')[1]}"
        self.app.con.execute(
            "INSERT INTO therapies(patient_id,date,note,next_visit,next_time) VALUES(?,?,?,?,?)",
            (self.patient["id"], get_date_iso(self.date), note, nv, nt))
        self.app.con.commit()
        self.destroy()
        self.app.reload_all()


class BloodDialog(Dialog):
    """Лабораториска анализа: две картички (хематолошка / биохемиска)."""

    def __init__(self, app, patient):
        super().__init__(app.root, "Нова лабораториска анализа")
        self.app, self.patient = app, patient

        top = ttk.Frame(self, padding=(14, 12, 14, 4))
        top.pack(fill="x")
        ttk.Label(top, text=f"{patient['pet_name']} - {patient['owner_name']} {patient['owner_surname']}",
                  font=(FONT, 12, "bold")).pack(side="left")
        self.date = make_date(top)
        self.date.pack(side="right")
        ttk.Label(top, text="Датум на примерок ").pack(side="right")
        self.species = tk.StringVar(value=patient["species"])
        cb = ttk.Combobox(top, textvariable=self.species, values=SPECIES, state="readonly", width=8)
        cb.pack(side="right", padx=(0, 20))
        ttk.Label(top, text="Вид (за референци) ").pack(side="right")
        cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_refs())

        self.nb = ttk.Notebook(self)
        self.nb.pack(padx=12, pady=6)
        self.rows = []
        for tab, (gkey, gtitle, _, params) in enumerate(GROUPS):
            fr = ttk.Frame(self.nb, padding=12)
            self.nb.add(fr, text=gtitle)
            per = (len(params) + 1) // 2
            for i, p in enumerate(params):
                block, r = divmod(i, per)
                c0 = block * 6
                ttk.Label(fr, text=p[0], width=31).grid(row=r, column=c0, sticky="w", pady=2)
                var = tk.StringVar()
                e = ttk.Entry(fr, textvariable=var, width=9)
                e.grid(row=r, column=c0 + 1, padx=4)
                ttk.Label(fr, text=unit_ui(p[1]), foreground=MUTED, width=8).grid(row=r, column=c0 + 2, sticky="w")
                ref = ttk.Label(fr, text="", foreground=MUTED, width=13)
                ref.grid(row=r, column=c0 + 3, sticky="w")
                fl = tk.Label(fr, text="", font=(FONT, 12, "bold"), width=2, bg=BG)
                fl.grid(row=r, column=c0 + 4)
                fr.grid_columnconfigure(c0 + 5, minsize=28)
                row = dict(group=gkey, param=p, var=var, ref=ref, flag=fl, entry=e, tab=tab)
                var.trace_add("write", lambda *a, rw=row: self.update_flag(rw))
                self.rows.append(row)
        for idx, row in enumerate(self.rows):
            row["entry"].bind("<Return>", lambda ev, i=idx: self.goto(i + 1))
        self.refresh_refs()
        self.rows[0]["entry"].focus_set()

        nf = ttk.Frame(self, padding=(14, 4))
        nf.pack(fill="x")
        ttk.Label(nf, text="Забелешка").pack(side="left")
        self.note = tk.StringVar()
        ttk.Entry(nf, textvariable=self.note, width=70).pack(side="left", padx=8, fill="x", expand=True)
        self.buttons("Зачувај PDF", self.save)
        self.show()

    def goto(self, i):
        row = self.rows[i % len(self.rows)]
        self.nb.select(row["tab"])
        row["entry"].focus_set()

    def refresh_refs(self):
        for row in self.rows:
            ref = get_ref(row["param"], self.species.get())
            row["ref"].config(text=ref[0] if ref else "—")
            self.update_flag(row)

    def update_flag(self, row):
        try:
            v = parse_num(row["var"].get())
        except ValueError:
            v = None
        f = compute_flag(v, get_ref(row["param"], self.species.get()))
        row["flag"].config(text=f, fg=RED if f == UP else BLUE)

    def save(self):
        key = self.species.get()
        out = {"hem": [], "bio": []}
        for row in self.rows:
            raw = row["var"].get().strip()
            if not raw:
                continue
            try:
                v = parse_num(raw)
            except ValueError:
                messagebox.showwarning("Внимание", f"Невалидна вредност за {row['param'][0]}: {raw}", parent=self)
                return
            ref = get_ref(row["param"], key)
            out[row["group"]].append(dict(name=row["param"][0], unit=row["param"][1],
                                          value=raw.replace(",", "."), flag=compute_flag(v, ref),
                                          ref=ref[0] if ref else ""))
        if not out["hem"] and not out["bio"]:
            messagebox.showwarning("Внимание", "Внеси барем еден параметар.", parent=self)
            return
        d = get_date_iso(self.date)
        pat = dict(self.patient)
        pat["species"] = key
        owner = safe_name(f"{pat['owner_name']}_{pat['owner_surname']}")
        path = unique_path(BLOOD_DIR, f"{owner}_крвна_слика_{d}.pdf")
        try:
            make_lab_pdf(pat, d, out["hem"], out["bio"], self.note.get(), path)
        except Exception as e:
            messagebox.showerror("Грешка", f"PDF не може да се создаде:\n{e}", parent=self)
            return
        self.app.con.execute("INSERT INTO bloodtests(patient_id,date,data,path) VALUES(?,?,?,?)",
                             (self.patient["id"], d, json.dumps(out, ensure_ascii=False), path))
        self.app.con.commit()
        self.destroy()
        self.app.reload_all()
        open_file(path)


class ReportDialog(Dialog):
    SECTIONS = [("Анамнеза", 4), ("Клинички преглед", 4), ("Дијагноза", 3),
                ("Терапија / препорака", 4), ("Забелешка", 3)]

    def __init__(self, app, patient):
        super().__init__(app.root, "Нов извештај")
        self.app, self.patient = app, patient
        top = ttk.Frame(self, padding=(14, 12, 14, 4))
        top.pack(fill="x")
        ttk.Label(top, text=f"{patient['pet_name']} ({patient['species']}) - "
                            f"{patient['owner_name']} {patient['owner_surname']}",
                  font=(FONT, 12, "bold")).pack(side="left")
        self.date = make_date(top)
        self.date.pack(side="right")
        ttk.Label(top, text="Датум ").pack(side="right")
        f = ttk.Frame(self, padding=(14, 4))
        f.pack()
        self.texts = []
        for i, (title, h) in enumerate(self.SECTIONS):
            ttk.Label(f, text=title, font=(FONT, 11, "bold"), foreground=PRIMARY).grid(
                row=2 * i, column=0, sticky="w", pady=(8, 0))
            t = tk.Text(f, width=80, height=h, font=(FONT, 11), wrap="word", relief="solid", bd=1)
            t.grid(row=2 * i + 1, column=0)
            self.texts.append((title, t))
        self.buttons("Зачувај PDF", self.save)
        self.show()

    def save(self):
        sections = [(t, w.get("1.0", "end").strip()) for t, w in self.texts]
        if not any(b for _, b in sections):
            messagebox.showwarning("Внимание", "Извештајот е празен.", parent=self)
            return
        d = get_date_iso(self.date)
        pat = dict(self.patient)
        owner = safe_name(f"{pat['owner_name']}_{pat['owner_surname']}")
        path = unique_path(REPORT_DIR, f"{owner}_извештај_{d}.pdf")
        try:
            make_report_pdf(pat, d, sections, path)
        except Exception as e:
            messagebox.showerror("Грешка", f"PDF не може да се создаде:\n{e}", parent=self)
            return
        self.app.con.execute("INSERT INTO reports(patient_id,date,data,path) VALUES(?,?,?,?)",
                             (self.patient["id"], d, json.dumps(sections, ensure_ascii=False), path))
        self.app.con.commit()
        self.destroy()
        self.app.reload_all()
        open_file(path)


# ----------------------------------------------------------------------------
# Главен прозорец
# ----------------------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.con = sqlite3.connect(DB_PATH)
        self.con.row_factory = sqlite3.Row
        self.con.execute("PRAGMA foreign_keys=ON")
        self.con.executescript(SCHEMA)
        migrate(self.con)
        self.current_id = None

        root.title(CLINIC)
        root.configure(bg=BG)
        root.minsize(1100, 700)
        try:
            root.state("zoomed")
        except tk.TclError:
            root.geometry("1400x850")
        self.setup_style()
        self.build_header()
        self.build_main()
        self.build_calendar()
        root.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(2, weight=1)
        self.refresh_list()
        self.refresh_calendar()
        self.tick()
        if not FONTS_OK:
            messagebox.showwarning("Фонт", "Не е пронајден фонт со кирилица за PDF (Arial/DejaVu). "
                                           "PDF-овите може да не работат.")

    # ---------- стил ----------
    def setup_style(self):
        st = ttk.Style()
        st.theme_use("clam")
        st.configure(".", font=(FONT, 11), background=BG, foreground=TEXT)
        st.configure("TLabelframe", background=BG, bordercolor="#CBD5E1")
        st.configure("TLabelframe.Label", background=BG, foreground=PRIMARY, font=(FONT, 11, "bold"))
        st.configure("Treeview", rowheight=30, font=(FONT, 11), background="white", fieldbackground="white")
        st.configure("Treeview.Heading", font=(FONT, 11, "bold"), background=PRIMARY, foreground="white",
                     padding=6)
        st.map("Treeview.Heading", background=[("active", DARK)])
        st.map("Treeview", background=[("selected", ACCENT)], foreground=[("selected", "white")])
        st.configure("Accent.TButton", background=PRIMARY, foreground="white", font=(FONT, 11, "bold"),
                     padding=(14, 7), borderwidth=0)
        st.map("Accent.TButton", background=[("active", DARK)])
        st.configure("Danger.TButton", background=RED, foreground="white", font=(FONT, 10, "bold"),
                     padding=(10, 6), borderwidth=0)
        st.map("Danger.TButton", background=[("active", "#991B1B")])
        st.configure("TButton", padding=(10, 6))
        st.configure("TNotebook", background=BG, borderwidth=0)
        st.configure("TNotebook.Tab", padding=(18, 9), font=(FONT, 11, "bold"))
        st.map("TNotebook.Tab", background=[("selected", "white")], foreground=[("selected", PRIMARY)])

    # ---------- заглавие ----------
    def build_header(self):
        h = tk.Frame(self.root, bg=DARK)
        h.grid(row=0, column=0, sticky="ew")
        try:
            self.logo = tk.PhotoImage(data=LOGO_B64).subsample(4, 4)
            tk.Label(h, image=self.logo, bg=DARK).pack(side="left", padx=(24, 12), pady=8)
        except tk.TclError:
            tk.Label(h, text="🐾", font=(FONT, 30), bg=DARK, fg=ACCENT).pack(side="left", padx=(24, 8), pady=10)
        box = tk.Frame(h, bg=DARK)
        box.pack(side="left")
        tk.Label(box, text=CLINIC, font=(FONT, 24, "bold"), bg=DARK, fg="white").pack(anchor="w")
        tk.Label(box, text="Систем за водење пациенти, анализи и извештаи", font=(FONT, 11),
                 bg=DARK, fg="#99F6E4").pack(anchor="w")
        self.clock = tk.Label(h, text="", font=(FONT, 14, "bold"), bg=DARK, fg="white", justify="right")
        self.clock.pack(side="right", padx=28)
        tk.Frame(self.root, bg=ACCENT, height=4).grid(row=1, column=0, sticky="ew")

    def tick(self):
        n = datetime.now()
        self.clock.config(text=f"{DAYS[n.weekday()]}, {n.strftime('%d.%m.%Y')}\n{n.strftime('%H:%M:%S')}")
        if n.second == 0:
            self.refresh_calendar()
        self.root.after(1000, self.tick)

    # ---------- главен дел ----------
    def build_main(self):
        pw = ttk.PanedWindow(self.root, orient="horizontal")
        pw.grid(row=2, column=0, sticky="nsew", padx=12, pady=10)

        left = ttk.Frame(pw, padding=(0, 0, 8, 0))
        pw.add(left, weight=1)
        ttk.Label(left, text="🔍 Пребарување (сопственик, милениче или датум)",
                  font=(FONT, 11, "bold"), foreground=PRIMARY).pack(anchor="w")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self.refresh_list())
        ttk.Entry(left, textvariable=self.search_var, font=(FONT, 13)).pack(fill="x", pady=(4, 2), ipady=4)
        ttk.Label(left, text="пр. Петровски, Мишко, 07.11.2026 или 11.2026", foreground=MUTED,
                  font=(FONT, 9)).pack(anchor="w", pady=(0, 6))
        ttk.Button(left, text="＋  Нов пациент", style="Accent.TButton",
                   command=lambda: PatientDialog(self)).pack(fill="x", pady=(0, 8))
        cols = ("pet", "owner", "species")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse")
        for c, t, w in (("pet", "Милениче", 130), ("owner", "Сопственик", 170), ("species", "Вид", 70)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w)
        sb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        right = ttk.Frame(pw, padding=(8, 0, 0, 0))
        pw.add(right, weight=3)
        self.info = tk.Frame(right, bg="white", highlightbackground="#CBD5E1", highlightthickness=1)
        self.info.pack(fill="x", pady=(0, 8))
        self.info_title = tk.Label(self.info, text="Избери или додади пациент", font=(FONT, 18, "bold"),
                                   bg="white", fg=DARK, anchor="w")
        self.info_title.pack(side="top", anchor="w", padx=16, pady=(10, 0))
        self.info_sub = tk.Label(self.info, text="", font=(FONT, 11), bg="white", fg=MUTED, anchor="w")
        self.info_sub.pack(side="top", anchor="w", padx=16, pady=(0, 10))
        bf = tk.Frame(self.info, bg="white")
        bf.place(relx=1.0, rely=0.5, anchor="e", x=-12)
        ttk.Button(bf, text="✎ Уреди", command=self.edit_patient).pack(side="left", padx=4)
        ttk.Button(bf, text="🗑 Избриши", style="Danger.TButton", command=self.delete_patient).pack(side="left", padx=4)

        nb = ttk.Notebook(right)
        nb.pack(fill="both", expand=True)

        t1 = ttk.Frame(nb, padding=8)
        nb.add(t1, text="💊  Терапии и забелешки")
        bar = ttk.Frame(t1)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="＋ Нова терапија / забелешка", style="Accent.TButton",
                   command=lambda: self.need_patient(TherapyDialog)).pack(side="left")
        ttk.Button(bar, text="Избриши избрано", command=lambda: self.delete_row("therapies", self.t_tree)).pack(
            side="right")
        self.t_tree = self.make_tree(t1, [("date", "Датум", 100), ("note", "Терапија / забелешка", 520),
                                          ("next", "Следна посета", 150)])
        self.t_tree.bind("<Double-1>", self.show_therapy)

        t2 = ttk.Frame(nb, padding=8)
        nb.add(t2, text="🩸  Крвни слики")
        bar = ttk.Frame(t2)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="＋ Нова крвна слика", style="Accent.TButton",
                   command=lambda: self.need_patient(BloodDialog)).pack(side="left")
        ttk.Label(bar, text="  (двоен клик отвора PDF)", foreground=MUTED).pack(side="left")
        ttk.Button(bar, text="Избриши избрано", command=lambda: self.delete_row("bloodtests", self.b_tree)).pack(
            side="right")
        self.b_tree = self.make_tree(t2, [("date", "Датум", 110), ("file", "PDF датотека", 600)])
        self.b_tree.bind("<Double-1>", lambda e: self.open_pdf("bloodtests", self.b_tree))

        t3 = ttk.Frame(nb, padding=8)
        nb.add(t3, text="📄  Извештаи")
        bar = ttk.Frame(t3)
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="＋ Нов извештај", style="Accent.TButton",
                   command=lambda: self.need_patient(ReportDialog)).pack(side="left")
        ttk.Label(bar, text="  (двоен клик отвора PDF)", foreground=MUTED).pack(side="left")
        ttk.Button(bar, text="Избриши избрано", command=lambda: self.delete_row("reports", self.r_tree)).pack(
            side="right")
        self.r_tree = self.make_tree(t3, [("date", "Датум", 110), ("file", "PDF датотека", 600)])
        self.r_tree.bind("<Double-1>", lambda e: self.open_pdf("reports", self.r_tree))

    def make_tree(self, parent, cols):
        tr = ttk.Treeview(parent, columns=[c[0] for c in cols], show="headings", selectmode="browse")
        for c, t, w in cols:
            tr.heading(c, text=t)
            tr.column(c, width=w, anchor="w")
        sb = ttk.Scrollbar(parent, orient="vertical", command=tr.yview)
        tr.configure(yscrollcommand=sb.set)
        tr.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        return tr

    # ---------- календар ----------
    def build_calendar(self):
        outer = tk.Frame(self.root, bg=BG)
        outer.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 12))
        tk.Label(outer, text="📅  Закажани посети - следни 10 дена", font=(FONT, 12, "bold"),
                 bg=BG, fg=PRIMARY).pack(anchor="w", pady=(0, 4))
        self.cal = tk.Frame(outer, bg=BG)
        self.cal.pack(fill="x")
        for i in range(10):
            self.cal.grid_columnconfigure(i, weight=1, uniform="d")

    def refresh_calendar(self):
        for w in self.cal.winfo_children():
            w.destroy()
        today = date.today()
        end = today + timedelta(days=9)
        rows = self.con.execute(
            """SELECT t.next_visit, t.next_time, p.id AS pid, p.pet_name, p.owner_surname
               FROM therapies t JOIN patients p ON p.id=t.patient_id
               WHERE t.next_visit BETWEEN ? AND ? ORDER BY t.next_visit, t.next_time""",
            (today.isoformat(), end.isoformat())).fetchall()
        by_day = {}
        for r in rows:
            by_day.setdefault(r["next_visit"], []).append(r)
        for i in range(10):
            d = today + timedelta(days=i)
            is_today = i == 0
            card = tk.Frame(self.cal, bg="white", highlightthickness=2 if is_today else 1,
                            highlightbackground=ACCENT if is_today else "#CBD5E1")
            card.grid(row=0, column=i, sticky="nsew", padx=3)
            tk.Label(card, text=f"{DAYS_SHORT[d.weekday()]}  {d.strftime('%d.%m.')}",
                     font=(FONT, 11, "bold"), bg=PRIMARY if is_today else "#E2E8F0",
                     fg="white" if is_today else TEXT, pady=3).pack(fill="x")
            items = by_day.get(d.isoformat(), [])
            body = tk.Frame(card, bg="white", height=120)
            body.pack(fill="both", expand=True)
            body.pack_propagate(False)
            if not items:
                tk.Label(body, text="—", bg="white", fg="#CBD5E1").pack(pady=8)
            for r in items[:5]:
                txt = f"{r['next_time'] or '--:--'}  {r['pet_name']}\n{r['owner_surname']}"
                lb = tk.Label(body, text=txt, font=(FONT, 9, "bold"), bg="#CCFBF1", fg=DARK, anchor="w",
                              justify="left", cursor="hand2", padx=4, pady=1)
                lb.pack(fill="x", padx=3, pady=1)
                lb.bind("<Button-1>", lambda e, pid=r["pid"]: self.select_patient(pid))
            if len(items) > 5:
                tk.Label(body, text=f"+ уште {len(items) - 5}", font=(FONT, 9), bg="white", fg=MUTED).pack()

    # ---------- листа и пребарување ----------
    def refresh_list(self):
        q = self.search_var.get().strip().casefold()
        pts = self.con.execute(
            "SELECT * FROM patients ORDER BY owner_surname COLLATE NOCASE, owner_name COLLATE NOCASE").fetchall()
        dates = {}
        if q:
            for tbl, col in (("therapies", "date"), ("therapies", "next_visit"),
                             ("bloodtests", "date"), ("reports", "date")):
                for r in self.con.execute(f"SELECT patient_id, {col} AS d FROM {tbl} WHERE {col} != ''"):
                    dates.setdefault(r["patient_id"], set()).add(to_disp(r["d"]))
        self.tree.delete(*self.tree.get_children())
        for p in pts:
            hay = f"{p['owner_name']} {p['owner_surname']} {p['owner_surname']} {p['owner_name']} {p['pet_name']}".casefold()
            if q and not (q in hay or any(q in d for d in dates.get(p["id"], ()))):
                continue
            self.tree.insert("", "end", iid=str(p["id"]),
                             values=(p["pet_name"], f"{p['owner_name']} {p['owner_surname']}", p["species"]))
        if self.current_id and self.tree.exists(str(self.current_id)):
            self.tree.selection_set(str(self.current_id))

    def on_select(self, _=None):
        sel = self.tree.selection()
        if sel:
            self.current_id = int(sel[0])
            self.load_patient()

    def select_patient(self, pid):
        self.current_id = pid
        self.search_var.set("")
        self.refresh_list()
        if self.tree.exists(str(pid)):
            self.tree.selection_set(str(pid))
            self.tree.see(str(pid))
        self.load_patient()
        self.refresh_calendar()

    def get_patient(self):
        if not self.current_id:
            return None
        return self.con.execute("SELECT * FROM patients WHERE id=?", (self.current_id,)).fetchone()

    def load_patient(self):
        p = self.get_patient()
        if not p:
            self.info_title.config(text="Избери или додади пациент")
            self.info_sub.config(text="")
            for t in (self.t_tree, self.b_tree, self.r_tree):
                t.delete(*t.get_children())
            return
        icon = "🐕" if p["species"] == "Куче" else "🐈"
        self.info_title.config(text=f"{icon}  {p['pet_name']}  ·  {p['species']}")
        parts = [f"Сопственик: {p['owner_name']} {p['owner_surname']}"]
        for label, key in (("Тел", "phone"), ("Раса", "breed"), ("Старост", "age"),
                           ("Пол", "sex"), ("Микрочип", "microchip")):
            if p[key]:
                parts.append(f"{label}: {p[key]}")
        self.info_sub.config(text="   ·   ".join(parts))

        self.t_tree.delete(*self.t_tree.get_children())
        for r in self.con.execute("SELECT * FROM therapies WHERE patient_id=? ORDER BY date DESC, id DESC",
                                  (p["id"],)):
            nxt = (to_disp(r["next_visit"]) + (" " + r["next_time"] if r["next_time"] else "")) if r["next_visit"] else ""
            note = r["note"].replace("\n", "  ⏎  ")
            self.t_tree.insert("", "end", iid=str(r["id"]), values=(to_disp(r["date"]), note, nxt))
        for tbl, tree in (("bloodtests", self.b_tree), ("reports", self.r_tree)):
            tree.delete(*tree.get_children())
            for r in self.con.execute(f"SELECT * FROM {tbl} WHERE patient_id=? ORDER BY date DESC, id DESC",
                                      (p["id"],)):
                tree.insert("", "end", iid=str(r["id"]), values=(to_disp(r["date"]), os.path.basename(r["path"])))

    def reload_all(self):
        self.refresh_list()
        self.load_patient()
        self.refresh_calendar()

    # ---------- акции ----------
    def need_patient(self, dialog_cls):
        p = self.get_patient()
        if not p:
            messagebox.showinfo("Пациент", "Прво избери или додади пациент.")
            return
        dialog_cls(self, p)

    def edit_patient(self):
        p = self.get_patient()
        if p:
            PatientDialog(self, p)

    def delete_patient(self):
        p = self.get_patient()
        if not p:
            return
        if messagebox.askyesno("Бришење", f"Да се избрише пациентот {p['pet_name']} "
                                         f"({p['owner_name']} {p['owner_surname']}) и сите негови записи?\n\n"
                                         "(PDF датотеките во папките остануваат.)"):
            self.con.execute("DELETE FROM patients WHERE id=?", (p["id"],))
            self.con.commit()
            self.current_id = None
            self.reload_all()

    def delete_row(self, table, tree):
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Бришење", "Прво избери ред.")
            return
        if messagebox.askyesno("Бришење", "Да се избрише избраниот запис?\n(PDF датотеката останува во папката.)"):
            self.con.execute(f"DELETE FROM {table} WHERE id=?", (int(sel[0]),))
            self.con.commit()
            self.reload_all()

    def open_pdf(self, table, tree):
        sel = tree.selection()
        if not sel:
            return
        r = self.con.execute(f"SELECT path FROM {table} WHERE id=?", (int(sel[0]),)).fetchone()
        if r:
            open_file(r["path"])

    def show_therapy(self, _=None):
        sel = self.t_tree.selection()
        if not sel:
            return
        r = self.con.execute("SELECT * FROM therapies WHERE id=?", (int(sel[0]),)).fetchone()
        if not r:
            return
        d = Dialog(self.root, f"Забелешка од {to_disp(r['date'])}")
        t = tk.Text(d, width=70, height=12, font=(FONT, 11), wrap="word", padx=10, pady=10)
        t.pack(padx=12, pady=12)
        t.insert("1.0", r["note"])
        t.config(state="disabled")
        if r["next_visit"]:
            ttk.Label(d, text=f"Следна посета: {to_disp(r['next_visit'])} {r['next_time']}",
                      font=(FONT, 11, "bold"), foreground=PRIMARY).pack(pady=(0, 8))
        ttk.Button(d, text="Затвори", command=d.destroy).pack(pady=(0, 12))
        d.show()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()