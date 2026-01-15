# app.py
# Reportes de Rodaje — versión “pro” lista para .app (PyInstaller)
# - Fonts (assets) se leen desde el bundle via resource_path()
# - reports/exports (datos) se guardan SIEMPRE en Documents/ReportesRodaje/ (escritura segura)
# - PDF: tabla paginada, celdas con wrap (nunca se pegan letras), badges solo color (sin texto)
# - Reporte: editor con bold/italic/bullets + PDF paginado + soporte unicode/emoji (best-effort)

import json
import os
import sys
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# =========================
#  RUTAS “PRO” (CRÍTICO)
# =========================
def resource_path(relative_path: str) -> str:
    """
    Ruta correcta tanto en dev como dentro del .app (PyInstaller).
    SOLO para leer assets que van dentro del bundle.
    """
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def user_data_dir(app_folder_name="ReportesRodaje") -> str:
    """
    Carpeta segura para ESCRIBIR en macOS/Windows:
    ~/Documents/ReportesRodaje/
    """
    base = Path.home() / "Documents" / app_folder_name
    return str(base)


APP_TITLE = "Reportes de Rodaje"

# Assets (solo lectura dentro del bundle)
FONTS_DIR = resource_path("fonts")

# Datos (lectura/escritura) — NO usar _MEIPASS
DATA_DIR = user_data_dir("ReportesRodaje")
REPORTS_DIR = os.path.join(DATA_DIR, "reports")
EXPORTS_DIR = os.path.join(DATA_DIR, "exports")


# =========================
#  CONFIG
# =========================
STATUS_OPTIONS = ["WORK IN PROGRESS", "DONE", "ON HOLD", "CANCELLED"]
PRIORITY_OPTIONS = ["BAJA", "MEDIA", "ALTA", "URGENTE"]

STATUS_COLORS = {
    "WORK IN PROGRESS": "#F4B400",
    "DONE": "#34A853",
    "ON HOLD": "#9AA0A6",
    "CANCELLED": "#EA4335",
}
PRIORITY_COLORS = {
    "BAJA": "#34A853",
    "MEDIA": "#F4B400",
    "ALTA": "#FB8C00",
    "URGENTE": "#EA4335",
}

BACKUP_OPTIONS = ["PENDIENTE", "COMPLETADO"]

# PDF: solo color (sin texto)
STATE_COLOR = {"COMPLETADO": "#34A853", "PENDIENTE": "#EA4335"}


# =========================
#  HELPERS
# =========================
def ensure_dirs():
    os.makedirs(REPORTS_DIR, exist_ok=True)
    os.makedirs(EXPORTS_DIR, exist_ok=True)


def valid_date_mmddyyyy(s: str) -> bool:
    try:
        datetime.strptime(s.strip(), "%m/%d/%Y")
        return True
    except Exception:
        return False


def mmddyyyy_to_iso(s: str) -> str:
    return datetime.strptime(s.strip(), "%m/%d/%Y").strftime("%Y-%m-%d")


def iso_to_mmddyyyy(s: str) -> str:
    return datetime.strptime(s.strip(), "%Y-%m-%d").strftime("%m/%d/%Y")


def hex_to_rgb01(hex_color: str):
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16) / 255.0
    g = int(h[2:4], 16) / 255.0
    b = int(h[4:6], 16) / 255.0
    return r, g, b


def folder_size_bytes(path: str) -> int:
    total = 0
    for root, _, files in os.walk(path):
        for f in files:
            fp = os.path.join(root, f)
            try:
                total += os.path.getsize(fp)
            except OSError:
                pass
    return total


def human_size(num_bytes: int) -> str:
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if num_bytes < 1024:
            return f"{num_bytes:.0f} {unit}" if unit == "B" else f"{num_bytes:.2f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.2f} PB"


class Badge(ttk.Frame):
    """Badge UI con color."""
    def __init__(self, parent, textvariable: tk.StringVar, color_map: dict, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.textvariable = textvariable
        self.color_map = color_map

        self.label = tk.Label(
            self,
            textvariable=self.textvariable,
            fg="white",
            bg="#777777",
            padx=10,
            pady=4,
            font=("Segoe UI", 9, "bold"),
        )
        self.label.pack(anchor="w")
        self.textvariable.trace_add("write", lambda *_: self.refresh())
        self.refresh()

    def refresh(self):
        val = (self.textvariable.get() or "").strip()
        bg = self.color_map.get(val, "#777777")
        self.label.configure(bg=bg)


def register_unicode_fonts():
    """
    Registra fuentes Unicode para PDF (mejor soporte de caracteres / emojis).
    Si faltan, usa Helvetica fallback.
    """
    try:
        regular = os.path.join(FONTS_DIR, "DejaVuSans.ttf")
        bold = os.path.join(FONTS_DIR, "DejaVuSans-Bold.ttf")
        italic = os.path.join(FONTS_DIR, "DejaVuSans-Oblique.ttf")
        bolditalic = os.path.join(FONTS_DIR, "DejaVuSans-BoldOblique.ttf")

        if all(os.path.exists(p) for p in [regular, bold, italic, bolditalic]):
            if "DJV" not in pdfmetrics.getRegisteredFontNames():
                pdfmetrics.registerFont(TTFont("DJV", regular))
                pdfmetrics.registerFont(TTFont("DJV-B", bold))
                pdfmetrics.registerFont(TTFont("DJV-I", italic))
                pdfmetrics.registerFont(TTFont("DJV-BI", bolditalic))
            return True
    except Exception:
        pass
    return False


def safe_draw_string(pdf_canvas, x, y, s, font_name, font_size, color=colors.white):
    """
    Dibuja texto sin romper por caracteres raros.
    Si algo no puede renderizar, lo reemplaza por '□'.
    """
    if s is None:
        s = ""
    try:
        pdf_canvas.setFont(font_name, font_size)
        pdf_canvas.setFillColor(color)
        pdf_canvas.drawString(x, y, s)
        return
    except Exception:
        cleaned = "".join(ch if ord(ch) < 0xD800 else "□" for ch in s)
        cleaned = cleaned.encode("utf-8", "replace").decode("utf-8")
        cleaned = "".join(ch if ch.isprintable() else "□" for ch in cleaned)
        pdf_canvas.setFont(font_name, font_size)
        pdf_canvas.setFillColor(color)
        pdf_canvas.drawString(x, y, cleaned)


# =========================
#  EDITOR RICH TEXT
# =========================
class ReportEditor(ttk.Frame):
    """
    Editor simple con:
    - Bold / Italic (selección o línea actual)
    - Bullets (agrega "• " al inicio de cada línea seleccionada)
    Guarda payload:
      {"text": "...", "bold_ranges": [[start,end],...], "italic_ranges": [[start,end],...]}
    offsets son sobre el string text completo.
    """
    def __init__(self, parent):
        super().__init__(parent)

        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", pady=(0, 6))

        ttk.Button(toolbar, text="B", width=3, command=self.toggle_bold).pack(side="left")
        ttk.Button(toolbar, text="I", width=3, command=self.toggle_italic).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="•", width=3, command=self.add_bullets).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Quitar bullets", command=self.remove_bullets).pack(side="left", padx=(10, 0))

        self.text = tk.Text(self, wrap="word", height=14, undo=True)
        self.text.pack(fill="both", expand=True)

        self.text.tag_configure("bold", font=("Segoe UI", 10, "bold"))
        self.text.tag_configure("italic", font=("Segoe UI", 10, "italic"))
        self.text.configure(font=("Segoe UI", 10))

        hint = ttk.Label(self, text="Tip: seleccioná texto y presioná B/I. Bullets: seleccioná líneas y presioná •")
        hint.pack(anchor="w", pady=(6, 0))

    def _sel_or_current_line(self):
        try:
            start = self.text.index("sel.first")
            end = self.text.index("sel.last")
        except tk.TclError:
            start = self.text.index("insert linestart")
            end = self.text.index("insert lineend")
        return start, end

    def toggle_bold(self):
        start, end = self._sel_or_current_line()
        if "bold" in self.text.tag_names(start):
            self.text.tag_remove("bold", start, end)
        else:
            self.text.tag_add("bold", start, end)

    def toggle_italic(self):
        start, end = self._sel_or_current_line()
        if "italic" in self.text.tag_names(start):
            self.text.tag_remove("italic", start, end)
        else:
            self.text.tag_add("italic", start, end)

    def add_bullets(self):
        start, end = self._sel_or_current_line()
        line_start = self.text.index(f"{start} linestart")
        line_end = self.text.index(f"{end} lineend")
        cur = line_start
        while True:
            ls = self.text.index(f"{cur} linestart")
            le = self.text.index(f"{cur} lineend")
            if self.text.get(ls, f"{ls}+2c") != "• ":
                self.text.insert(ls, "• ")
            if self.text.compare(le, ">=", line_end):
                break
            cur = self.text.index(f"{le}+1c")

    def remove_bullets(self):
        start, end = self._sel_or_current_line()
        line_start = self.text.index(f"{start} linestart")
        line_end = self.text.index(f"{end} lineend")
        cur = line_start
        while True:
            ls = self.text.index(f"{cur} linestart")
            if self.text.get(ls, f"{ls}+2c") == "• ":
                self.text.delete(ls, f"{ls}+2c")
            le = self.text.index(f"{cur} lineend")
            if self.text.compare(le, ">=", line_end):
                break
            cur = self.text.index(f"{le}+1c")

    def _index_to_offset(self, index_str: str) -> int:
        return int(self.text.count("1.0", index_str, "chars")[0])

    def _offset_to_index(self, offset: int) -> str:
        return f"1.0+{offset}c"

    def get_payload(self) -> dict:
        txt = self.text.get("1.0", "end-1c")

        bold_ranges = []
        italic_ranges = []

        def collect(tag_name, out_list):
            ranges = self.text.tag_ranges(tag_name)
            for i in range(0, len(ranges), 2):
                s = ranges[i]
                e = ranges[i + 1]
                out_list.append([self._index_to_offset(str(s)), self._index_to_offset(str(e))])

        collect("bold", bold_ranges)
        collect("italic", italic_ranges)

        return {"text": txt, "bold_ranges": bold_ranges, "italic_ranges": italic_ranges}

    def set_payload(self, payload: dict):
        self.text.delete("1.0", "end")
        self.text.insert("1.0", payload.get("text", ""))

        self.text.tag_remove("bold", "1.0", "end")
        self.text.tag_remove("italic", "1.0", "end")

        for s, e in payload.get("bold_ranges", []):
            self.text.tag_add("bold", self._offset_to_index(s), self._offset_to_index(e))
        for s, e in payload.get("italic_ranges", []):
            self.text.tag_add("italic", self._offset_to_index(s), self._offset_to_index(e))


# =========================
#  APP
# =========================
class ReportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x760")
        self.minsize(1120, 760)

        ensure_dirs()

        self.current_path = None
        self.report_data = self.default_report()

        self.container = ttk.Frame(self, padding=14)
        self.container.pack(fill="both", expand=True)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (StartFrame, EditorFrame):
            frame = F(parent=self.container, controller=self)
            self.frames[F.__name__] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("StartFrame")

    def default_report(self):
        return {
            "logo_path": "",
            "status_general": STATUS_OPTIONS[0],
            "proyecto": "",
            "fecha_inicio": "",
            "fecha_fin": "",
            "encargado": "",
            "etapa": "",
            "prioridad": PRIORITY_OPTIONS[1],
            "dias": [{"nombre": "Día 1", "rows": []}],
            "reporte_rich": {"text": "", "bold_ranges": [], "italic_ranges": []},
        }

    def show_frame(self, name: str):
        frame = self.frames[name]
        frame.tkraise()
        if name == "EditorFrame":
            frame.load_from_data(self.report_data)

    def new_report(self):
        self.current_path = None
        self.report_data = self.default_report()
        self.show_frame("EditorFrame")

    def open_report(self):
        path = filedialog.askopenfilename(
            title="Abrir reporte",
            initialdir=REPORTS_DIR,
            filetypes=[("Reporte JSON", "*.json")]
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            base = self.default_report()
            base.update(data)

            if "dias" not in base or not isinstance(base["dias"], list) or len(base["dias"]) == 0:
                base["dias"] = [{"nombre": "Día 1", "rows": []}]
            for d in base["dias"]:
                d.setdefault("nombre", "Día")
                d.setdefault("rows", [])
                for r in d["rows"]:
                    r.setdefault("backup_b", "PENDIENTE")
                    r.setdefault("backup_expo", "PENDIENTE")
                    r.setdefault("proxies", "PENDIENTE")

            base.setdefault("reporte_rich", {"text": "", "bold_ranges": [], "italic_ranges": []})

            self.report_data = base
            self.current_path = path
            self.show_frame("EditorFrame")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el reporte.\n\n{e}")

    def save_report(self, data: dict):
        if not self.current_path:
            default_name = (data.get("proyecto") or "reporte").strip().replace(" ", "_") or "reporte"
            path = filedialog.asksaveasfilename(
                title="Guardar reporte",
                defaultextension=".json",
                initialdir=REPORTS_DIR,
                initialfile=f"{default_name}.json",
                filetypes=[("Reporte JSON", "*.json")],
            )
            if not path:
                return None
            self.current_path = path

        try:
            with open(self.current_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return self.current_path
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar.\n\n{e}")
            return None

    def export_pdf(self, data: dict):
        fi_iso = (data.get("fecha_inicio") or "").strip()
        ff_iso = (data.get("fecha_fin") or "").strip()

        def display_date(iso_s: str) -> str:
            if not iso_s:
                return ""
            try:
                return iso_to_mmddyyyy(iso_s)
            except Exception:
                return iso_s

        fi_display = display_date(fi_iso)
        ff_display = display_date(ff_iso)

        proyecto = (data.get("proyecto") or "reporte").strip().replace(" ", "_") or "reporte"
        out_path = os.path.join(EXPORTS_DIR, f"{proyecto}.pdf")

        status = data.get("status_general", "")
        prio = data.get("prioridad", "")
        status_hex = STATUS_COLORS.get(status, "#777777")
        prio_hex = PRIORITY_COLORS.get(prio, "#777777")

        # ----- PALETA DARK -----
        BG = colors.HexColor("#0B0F14")
        CARD = colors.HexColor("#121824")
        BORDER = colors.HexColor("#263042")
        DIVIDER = colors.HexColor("#243042")
        LABEL = colors.HexColor("#C7D0DD")
        VALUE = colors.HexColor("#FFFFFF")
        TITLE = colors.HexColor("#FFFFFF")

        def hex_to_rl(hex_color: str):
            r, g, b = hex_to_rgb01(hex_color)
            return colors.Color(r, g, b)

        # ---- PDF helpers ----
        def wrap_lines(pdf_canvas, text, font_name, font_size, max_width):
            """
            Wrap robusto para que NUNCA se peguen letras:
            - separa por espacios
            - si una palabra es muy larga, la corta por caracteres
            """
            s = str(text or "").replace("\t", " ").strip()
            if not s:
                return [""]

            pdf_canvas.setFont(font_name, font_size)
            words = s.split(" ")
            lines = []
            cur = ""

            def flush():
                nonlocal cur
                if cur.strip():
                    lines.append(cur.strip())
                cur = ""

            def add_token(tok):
                nonlocal cur
                test = (cur + " " + tok).strip()
                if pdf_canvas.stringWidth(test, font_name, font_size) <= max_width:
                    cur = test
                    return

                if cur:
                    flush()

                if pdf_canvas.stringWidth(tok, font_name, font_size) <= max_width:
                    cur = tok
                    return

                # cortar palabra por caracteres
                buf = ""
                for ch in tok:
                    candidate = buf + ch
                    if pdf_canvas.stringWidth(candidate, font_name, font_size) <= max_width:
                        buf = candidate
                    else:
                        if buf:
                            lines.append(buf)
                        buf = ch
                if buf:
                    lines.append(buf)

            for w in words:
                if not w:
                    continue
                add_token(w)

            flush()
            return lines if lines else [""]

        def draw_lines(pdf_canvas, x, y_top, lines, font_name, font_size, line_h, color):
            y = y_top
            for ln in lines:
                safe_draw_string(pdf_canvas, x, y, ln, font_name, font_size, color)
                y -= line_h

        def draw_badge_color_only(pdf_canvas, x, y, w, h, hexcol):
            r, g, b = hex_to_rgb01(hexcol)
            pdf_canvas.setFillColor(colors.Color(r, g, b))
            pdf_canvas.roundRect(x, y, w, h, 6, fill=1, stroke=0)

        # ---- Reporte rich text paginado ----
        def render_report_pages(pdf_canvas, payload, left_margin):
            text = (payload.get("text", "") or "")
            bold_ranges = payload.get("bold_ranges", []) or []
            italic_ranges = payload.get("italic_ranges", []) or []

            has_unicode = register_unicode_fonts()
            FONT_N = "DJV" if has_unicode else "Helvetica"
            FONT_B = "DJV-B" if has_unicode else "Helvetica-Bold"
            FONT_I = "DJV-I" if has_unicode else "Helvetica-Oblique"
            FONT_BI = "DJV-BI" if has_unicode else "Helvetica-BoldOblique"

            def style_for_offset(off):
                b = any(a <= off < b2 for a, b2 in bold_ranges)
                i = any(a <= off < b2 for a, b2 in italic_ranges)
                if b and i:
                    return "BI"
                if b:
                    return "B"
                if i:
                    return "I"
                return "N"

            def font_for_style(st):
                return {"N": FONT_N, "B": FONT_B, "I": FONT_I, "BI": FONT_BI}.get(st, FONT_N)

            def draw_rich_line(line_text, line_start_offset, x, y, fs):
                if line_text is None:
                    line_text = ""
                cx = x
                cur_run = ""
                cur_style = None

                def flush(run, st):
                    nonlocal cx
                    if not run:
                        return
                    fn = font_for_style(st)
                    safe_draw_string(pdf_canvas, cx, y, run, fn, fs, colors.white)
                    cx += pdf_canvas.stringWidth(run, fn, fs)

                for j, ch in enumerate(line_text):
                    off = line_start_offset + j
                    st = style_for_offset(off)
                    if cur_style is None:
                        cur_style = st
                        cur_run = ch
                    elif st == cur_style:
                        cur_run += ch
                    else:
                        flush(cur_run, cur_style)
                        cur_run = ch
                        cur_style = st
                flush(cur_run, cur_style)

            def wrap_text_with_offsets(start_offset, max_w, font_name, font_size):
                n = len(text)
                if start_offset >= n:
                    return None, None, n

                if text[start_offset] == "\n":
                    return "", start_offset, start_offset + 1

                nl = text.find("\n", start_offset)
                para_end = n if nl == -1 else nl
                para = text[start_offset:para_end]

                pdf_canvas.setFont(font_name, font_size)

                i = 0
                best_cut = 0
                last_space = -1
                while i < len(para):
                    if para[i] == " ":
                        last_space = i
                    candidate = para[: i + 1]
                    if pdf_canvas.stringWidth(candidate, font_name, font_size) <= max_w:
                        best_cut = i + 1
                        i += 1
                    else:
                        break

                if best_cut == len(para):
                    line = para
                    next_off = para_end
                    if nl != -1:
                        next_off += 1
                    return line, start_offset, next_off

                if last_space > 0:
                    line = para[:last_space]
                    next_off = start_offset + last_space + 1
                    while next_off < para_end and text[next_off] == " ":
                        next_off += 1
                    return line, start_offset, next_off

                cut = max(1, best_cut)
                line = para[:cut]
                next_off = start_offset + cut
                return line, start_offset, next_off

            width, height = LETTER
            card_x = left_margin
            card_w = width - 2 * left_margin
            card_y_top = height - 1.25 * inch
            card_h = height - 2.2 * inch

            text_x = card_x + 0.25 * inch
            text_y_top = card_y_top - 0.35 * inch
            max_w = card_w - 0.5 * inch
            max_h = card_h - 0.55 * inch

            fs = 10
            lh = 0.20 * inch
            start_off = 0
            page_num = 1

            def draw_page(pn):
                pdf_canvas.setFillColor(BG)
                pdf_canvas.rect(0, 0, width, height, fill=1, stroke=0)

                pdf_canvas.setFillColor(TITLE)
                pdf_canvas.setFont("Helvetica-Bold", 16)
                title = "Reporte" if pn == 1 else f"Reporte (cont. {pn})"
                pdf_canvas.drawString(left_margin, height - 0.9 * inch, title)

                pdf_canvas.setFillColor(CARD)
                pdf_canvas.roundRect(card_x, card_y_top - card_h, card_w, card_h, 12, fill=1, stroke=0)
                pdf_canvas.setStrokeColor(BORDER)
                pdf_canvas.setLineWidth(1)
                pdf_canvas.roundRect(card_x, card_y_top - card_h, card_w, card_h, 12, fill=0, stroke=1)

            if not text.strip():
                draw_page(1)
                safe_draw_string(pdf_canvas, text_x, text_y_top, "(sin texto)", FONT_N, 10, LABEL)
                pdf_canvas.showPage()
                return

            while start_off < len(text):
                draw_page(page_num)

                y = text_y_top
                used = 0.0
                while start_off < len(text) and used + lh <= max_h:
                    line, line_start, next_off = wrap_text_with_offsets(start_off, max_w, FONT_N, fs)
                    if line is None:
                        start_off = next_off
                        break
                    draw_rich_line(line, line_start, text_x, y, fs)
                    y -= lh
                    used += lh
                    start_off = next_off

                pdf_canvas.showPage()
                page_num += 1

        # ---- EXPORT PDF ----
        try:
            c = canvas.Canvas(out_path, pagesize=LETTER)
            width, height = LETTER
            left_margin = 0.85 * inch

            # ========== PORTADA ==========
            c.setFillColor(BG)
            c.rect(0, 0, width, height, fill=1, stroke=0)

            header_top = height - 0.65 * inch
            logo_path = data.get("logo_path", "")
            logo_size = 0.80 * inch

            block_x = left_margin
            block_top = header_top

            if logo_path and os.path.exists(logo_path):
                logo_y = block_top - logo_size
                c.drawImage(
                    logo_path, block_x, logo_y,
                    width=logo_size, height=logo_size,
                    preserveAspectRatio=True, mask="auto",
                )
            else:
                logo_y = block_top - logo_size

            title_y = logo_y - 0.18 * inch

            c.setFillColor(TITLE)
            c.setFont("Helvetica-Bold", 20)
            c.drawString(block_x, title_y, "RODAJE")

            header_y = title_y - 0.35 * inch

            card_x = left_margin
            card_w = width - 2 * left_margin
            card_y_top = header_y - 0.25 * inch
            row_h = 0.52 * inch
            rows = 7
            pad_y = 0.32 * inch
            card_h = rows * row_h + pad_y * 1.1

            c.setFillColor(CARD)
            c.roundRect(card_x, card_y_top - card_h, card_w, card_h, 12, fill=1, stroke=0)
            c.setStrokeColor(BORDER)
            c.setLineWidth(1)
            c.roundRect(card_x, card_y_top - card_h, card_w, card_h, 12, fill=0, stroke=1)

            label_w = 2.05 * inch
            label_x = card_x + 0.30 * inch
            value_x = card_x + label_w + 0.35 * inch

            items = [
                ("STATUS GENERAL", ("BADGE", status, status_hex)),
                ("PROYECTO", ("TEXT", data.get("proyecto", ""))),
                ("FECHA INICIO", ("TEXT", fi_display)),
                ("FECHA FIN", ("TEXT", ff_display)),
                ("ENCARGADO", ("TEXT", data.get("encargado", ""))),
                ("ETAPA", ("TEXT", data.get("etapa", ""))),
                ("PRIORIDAD", ("BADGE", prio, prio_hex)),
            ]

            y = card_y_top - 0.45 * inch
            for i, (label, valueinfo) in enumerate(items):
                if i > 0:
                    c.setStrokeColor(DIVIDER)
                    c.setLineWidth(1)
                    c.line(card_x + 0.25 * inch, y + 0.27 * inch, card_x + card_w - 0.25 * inch, y + 0.27 * inch)

                c.setFillColor(LABEL)
                c.setFont("Helvetica", 9)
                c.drawString(label_x, y, label)

                kind = valueinfo[0]
                if kind == "TEXT":
                    txt = str(valueinfo[1])
                    c.setFillColor(VALUE)
                    c.setFont("Helvetica-Bold", 10.5)
                    c.drawString(value_x, y, txt)
                else:
                    txt = str(valueinfo[1])
                    hexcol = valueinfo[2]
                    badge_fill = hex_to_rl(hexcol)
                    badge_h2 = 0.28 * inch
                    badge_w2 = max(1.25 * inch, (len(txt) * 0.11 + 0.70) * inch)
                    bx = value_x
                    by = y - 0.08 * inch
                    c.setFillColor(badge_fill)
                    c.roundRect(bx, by, badge_w2, badge_h2, 7, fill=1, stroke=0)
                    c.setFillColor(colors.white)
                    c.setFont("Helvetica-Bold", 9)
                    c.drawString(bx + 0.14 * inch, by + 0.08 * inch, txt)

                y -= row_h

            c.showPage()

            # ========== DÍAS ==========
            dias = data.get("dias", [])

            # headers cortos para evitar “mezcla”
            cols = [
                ("TARJETA", 2.15),
                ("BK B", 0.95),
                ("BK EXPO", 1.05),
                ("PROXIES", 1.05),
                ("FILES", 1.35),
                ("PESO", 0.95),
                ("DESGLOSE", 1.70),
                ("COMENTARIOS", 1.80),
            ]

            table_x = left_margin
            table_w = width - 2 * left_margin
            total_units = sum(u for _, u in cols)
            col_widths = [table_w * (u / total_units) for _, u in cols]

            th = 0.42 * inch
            pad_x = 0.14 * inch
            pad_y_cell = 0.10 * inch

            font_body = "Helvetica"
            font_bold = "Helvetica-Bold"
            fs_body = 8.4
            fs_tarjeta = 8.6
            line_h = 0.16 * inch
            badge_h = 0.28 * inch

            def draw_day_header(title_text: str, dia_nombre_real: str):
                c.setFillColor(BG)
                c.rect(0, 0, width, height, fill=1, stroke=0)

                c.setFillColor(TITLE)
                c.setFont("Helvetica-Bold", 16)
                c.drawString(left_margin, height - 0.9 * inch, title_text)

                c.setFillColor(LABEL)
                c.setFont("Helvetica-Bold", 11)
                c.drawString(left_margin, height - 1.25 * inch, f"REPORTE DE DATA {dia_nombre_real.upper()}")

            def draw_table_header(table_top):
                c.setFillColor(colors.HexColor("#0F1623"))
                c.roundRect(table_x, table_top - th - 0.05 * inch, table_w, th + 0.05 * inch, 10, fill=1, stroke=0)

                cx = table_x
                c.setFillColor(LABEL)
                c.setFont("Helvetica-Bold", 8)
                for (name, _u), wcol in zip(cols, col_widths):
                    c.drawString(cx + pad_x, table_top - th + 0.08 * inch, name)
                    cx += wcol

                cx_lines = table_x
                c.setStrokeColor(DIVIDER)
                c.setLineWidth(1)
                for wcol in col_widths[:-1]:
                    cx_lines += wcol
                    c.line(cx_lines, table_top - th - 0.02 * inch, cx_lines, table_top - 0.02 * inch)

            def compute_row_height(row_dict):
                tarjeta = row_dict.get("tarjeta", "")
                files_name = os.path.basename(row_dict.get("files_path", "")) if row_dict.get("files_path") else ""
                peso = row_dict.get("peso", "")
                desglose = row_dict.get("desglose", "")
                comentarios = row_dict.get("comentarios", "")

                lines_tarjeta = wrap_lines(c, tarjeta, font_bold, fs_tarjeta, (col_widths[0] - 2 * pad_x))
                lines_files = wrap_lines(c, files_name, font_body, fs_body, (col_widths[4] - 2 * pad_x))
                lines_peso = wrap_lines(c, peso, font_body, fs_body, (col_widths[5] - 2 * pad_x))
                lines_desg = wrap_lines(c, desglose, font_body, fs_body, (col_widths[6] - 2 * pad_x))
                lines_com = wrap_lines(c, comentarios, font_body, fs_body, (col_widths[7] - 2 * pad_x))

                max_lines = max(len(lines_tarjeta), len(lines_files), len(lines_peso), len(lines_desg), len(lines_com), 1)
                h_text = max_lines * line_h + 2 * pad_y_cell
                h_badge_min = badge_h + 2 * pad_y_cell
                return max(h_text, h_badge_min, 0.46 * inch)

            for d in dias:
                dia_nombre = d.get("nombre", "Día")
                rows_data = d.get("rows", [])

                table_top = height - 1.65 * inch
                bottom_margin = 0.9 * inch

                page_idx = 0
                row_i = 0
                total_rows = len(rows_data)

                if total_rows == 0:
                    draw_day_header(dia_nombre, dia_nombre)
                    table_h = th + 0.60 * inch
                    c.setFillColor(CARD)
                    c.roundRect(table_x, table_top - table_h, table_w, table_h, 10, fill=1, stroke=0)
                    c.setStrokeColor(BORDER)
                    c.roundRect(table_x, table_top - table_h, table_w, table_h, 10, fill=0, stroke=1)
                    draw_table_header(table_top)

                    c.setFillColor(LABEL)
                    c.setFont(font_body, 9)
                    c.drawString(table_x + pad_x, table_top - th - 0.35 * inch, "Sin filas.")
                    c.showPage()
                    continue

                while row_i < total_rows:
                    if page_idx == 0:
                        title_text = dia_nombre
                    elif page_idx == 1:
                        title_text = f"{dia_nombre} (cont.)"
                    else:
                        title_text = f"{dia_nombre} (cont. {page_idx})"

                    draw_day_header(title_text, dia_nombre)

                    y_cursor = table_top - th - 0.15 * inch
                    available = y_cursor - bottom_margin

                    chunk = []
                    chunk_heights = []
                    temp_i = row_i
                    used = 0.0
                    while temp_i < total_rows:
                        rh = compute_row_height(rows_data[temp_i])
                        if used + rh > available and chunk:
                            break
                        if rh > available and not chunk:
                            chunk.append(rows_data[temp_i])
                            chunk_heights.append(rh)
                            temp_i += 1
                            break
                        chunk.append(rows_data[temp_i])
                        chunk_heights.append(rh)
                        used += rh
                        temp_i += 1

                    table_h = th + used + 0.25 * inch
                    c.setFillColor(CARD)
                    c.roundRect(table_x, table_top - table_h, table_w, table_h, 10, fill=1, stroke=0)
                    c.setStrokeColor(BORDER)
                    c.roundRect(table_x, table_top - table_h, table_w, table_h, 10, fill=0, stroke=1)

                    draw_table_header(table_top)

                    yrow_top = table_top - th - 0.10 * inch

                    for row, row_hh in zip(chunk, chunk_heights):
                        c.setStrokeColor(DIVIDER)
                        c.setLineWidth(1)
                        c.line(table_x + 0.15 * inch, yrow_top, table_x + table_w - 0.15 * inch, yrow_top)

                        tarjeta_txt = row.get("tarjeta", "")
                        dropbox_url = (row.get("dropbox", "") or "").strip()

                        backup_b = (row.get("backup_b", "PENDIENTE") or "").strip().upper()
                        backup_expo = (row.get("backup_expo", "PENDIENTE") or "").strip().upper()
                        proxies = (row.get("proxies", "PENDIENTE") or "").strip().upper()

                        files_name = os.path.basename(row.get("files_path", "")) if row.get("files_path") else ""
                        peso = row.get("peso", "")
                        desglose = row.get("desglose", "")
                        comentarios = row.get("comentarios", "")

                        # vertical separators
                        cx_lines = table_x
                        c.setStrokeColor(DIVIDER)
                        c.setLineWidth(1)
                        for wcol in col_widths[:-1]:
                            cx_lines += wcol
                            c.line(cx_lines, yrow_top - row_hh + 0.08 * inch, cx_lines, yrow_top - 0.08 * inch)

                        cell_top_text_y = yrow_top - pad_y_cell - 0.05 * inch
                        cx = table_x

                        # TARJETA (wrap + hyperlink)
                        x0 = cx + pad_x
                        w0 = col_widths[0] - 2 * pad_x
                        lines_tarjeta = wrap_lines(c, tarjeta_txt, font_bold, fs_tarjeta, w0)
                        draw_lines(c, x0, cell_top_text_y, lines_tarjeta, font_bold, fs_tarjeta, line_h, VALUE)
                        if dropbox_url.startswith("http"):
                            block_h = max(1, len(lines_tarjeta)) * line_h
                            c.linkURL(dropbox_url, (x0, cell_top_text_y - block_h + 2, x0 + w0, cell_top_text_y + 10), relative=0)

                        # BK B (solo color)
                        cx += col_widths[0]
                        x1 = cx + pad_x
                        w1 = col_widths[1] - 2 * pad_x
                        badge_w1 = min(w1, 1.35 * inch)
                        bx1 = x1 + (w1 - badge_w1) / 2
                        by1 = (yrow_top - row_hh) + (row_hh - badge_h) / 2
                        draw_badge_color_only(c, bx1, by1, badge_w1, badge_h, STATE_COLOR.get(backup_b, "#9AA0A6"))

                        # BK EXPO (solo color)
                        cx += col_widths[1]
                        x2 = cx + pad_x
                        w2 = col_widths[2] - 2 * pad_x
                        badge_w2 = min(w2, 1.35 * inch)
                        bx2 = x2 + (w2 - badge_w2) / 2
                        by2 = (yrow_top - row_hh) + (row_hh - badge_h) / 2
                        draw_badge_color_only(c, bx2, by2, badge_w2, badge_h, STATE_COLOR.get(backup_expo, "#9AA0A6"))

                        # PROXIES (solo color)
                        cx += col_widths[2]
                        xP = cx + pad_x
                        wP = col_widths[3] - 2 * pad_x
                        badge_wP = min(wP, 1.35 * inch)
                        bxp = xP + (wP - badge_wP) / 2
                        byp = (yrow_top - row_hh) + (row_hh - badge_h) / 2
                        draw_badge_color_only(c, bxp, byp, badge_wP, badge_h, STATE_COLOR.get(proxies, "#9AA0A6"))

                        # FILES
                        cx += col_widths[3]
                        x3 = cx + pad_x
                        w3 = col_widths[4] - 2 * pad_x
                        lines_files = wrap_lines(c, files_name, font_body, fs_body, w3)
                        draw_lines(c, x3, cell_top_text_y, lines_files, font_body, fs_body, line_h, VALUE)

                        # PESO
                        cx += col_widths[4]
                        x4 = cx + pad_x
                        w4 = col_widths[5] - 2 * pad_x
                        lines_peso = wrap_lines(c, peso, font_body, fs_body, w4)
                        draw_lines(c, x4, cell_top_text_y, lines_peso, font_body, fs_body, line_h, VALUE)

                        # DESGLOSE
                        cx += col_widths[5]
                        x5 = cx + pad_x
                        w5 = col_widths[6] - 2 * pad_x
                        lines_desg = wrap_lines(c, desglose, font_body, fs_body, w5)
                        draw_lines(c, x5, cell_top_text_y, lines_desg, font_body, fs_body, line_h, VALUE)

                        # COMENTARIOS
                        cx += col_widths[6]
                        x6 = cx + pad_x
                        w6 = col_widths[7] - 2 * pad_x
                        lines_com = wrap_lines(c, comentarios, font_body, fs_body, w6)
                        draw_lines(c, x6, cell_top_text_y, lines_com, font_body, fs_body, line_h, VALUE)

                        yrow_top -= row_hh

                    c.setFillColor(LABEL)
                    c.setFont("Helvetica", 9)
                    c.drawRightString(width - left_margin, 0.55 * inch, f"Página {page_idx + 1}")

                    c.showPage()
                    row_i += len(chunk)
                    page_idx += 1

            # ========== REPORTE (paginado + unicode/emojis) ==========
            payload = data.get("reporte_rich", {"text": "", "bold_ranges": [], "italic_ranges": []})
            render_report_pages(c, payload, left_margin)

            c.save()
            messagebox.showinfo("PDF exportado", f"PDF generado en:\n{out_path}")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar el PDF.\n\n{e}")


# =========================
#  UI FRAMES
# =========================
class StartFrame(ttk.Frame):
    def __init__(self, parent, controller: ReportApp):
        super().__init__(parent)
        self.controller = controller

        ttk.Label(self, text="Reportes de Rodaje", font=("Segoe UI", 18, "bold")).pack(pady=(40, 16))
        ttk.Label(self, text="Elegí una opción para continuar.").pack(pady=(0, 22))

        btns = ttk.Frame(self)
        btns.pack()

        ttk.Button(btns, text="Nuevo reporte", width=22, command=controller.new_report)\
            .grid(row=0, column=0, padx=10, pady=6)
        ttk.Button(btns, text="Abrir reporte existente", width=22, command=controller.open_report)\
            .grid(row=0, column=1, padx=10, pady=6)


class EditorFrame(ttk.Frame):
    def __init__(self, parent, controller: ReportApp):
        super().__init__(parent)
        self.controller = controller

        self.logo_path = tk.StringVar()
        self.status_general = tk.StringVar()
        self.proyecto = tk.StringVar()
        self.fecha_inicio_ui = tk.StringVar()
        self.fecha_fin_ui = tk.StringVar()
        self.encargado = tk.StringVar()
        self.etapa = tk.StringVar()
        self.prioridad = tk.StringVar()

        header = ttk.Frame(self)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text="Editor de Reporte", font=("Segoe UI", 16, "bold")).pack(side="left")
        ttk.Button(header, text="Volver", command=lambda: controller.show_frame("StartFrame")).pack(side="right")

        top = ttk.Frame(self)
        top.pack(fill="x")

        card = ttk.LabelFrame(top, text="Datos del reporte", padding=14)
        card.pack(fill="x")
        card.grid_columnconfigure(0, weight=0)
        card.grid_columnconfigure(1, weight=1)

        def row(r, label, widget):
            ttk.Label(card, text=label).grid(row=r, column=0, sticky="w", pady=8, padx=(0, 12))
            widget.grid(row=r, column=1, sticky="ew", pady=8)

        logo_row = ttk.Frame(card)
        logo_row.grid_columnconfigure(0, weight=1)
        ttk.Entry(logo_row, textvariable=self.logo_path).grid(row=0, column=0, sticky="ew")
        ttk.Button(logo_row, text="Elegir…", command=self.pick_logo).grid(row=0, column=1, padx=(10, 0))
        row(0, "Logo", logo_row)

        status_row = ttk.Frame(card)
        status_row.grid_columnconfigure(0, weight=1)
        ttk.Combobox(status_row, textvariable=self.status_general, values=STATUS_OPTIONS, state="readonly")\
            .grid(row=0, column=0, sticky="ew")
        Badge(status_row, self.status_general, STATUS_COLORS).grid(row=0, column=1, padx=(10, 0), sticky="w")
        row(1, "Status general", status_row)

        row(2, "Proyecto", ttk.Entry(card, textvariable=self.proyecto))

        dates = ttk.Frame(card)
        ttk.Label(dates, text="Inicio:").grid(row=0, column=0, sticky="w")
        ttk.Entry(dates, width=14, textvariable=self.fecha_inicio_ui).grid(row=0, column=1, padx=(6, 16), sticky="w")
        ttk.Label(dates, text="Fin:").grid(row=0, column=2, sticky="w")
        ttk.Entry(dates, width=14, textvariable=self.fecha_fin_ui).grid(row=0, column=3, padx=(6, 0), sticky="w")
        row(3, "Fechas (MM/DD/YYYY)", dates)

        row(4, "Encargado", ttk.Entry(card, textvariable=self.encargado))
        row(5, "Etapa", ttk.Entry(card, textvariable=self.etapa))

        prio_row = ttk.Frame(card)
        prio_row.grid_columnconfigure(0, weight=1)
        ttk.Combobox(prio_row, textvariable=self.prioridad, values=PRIORITY_OPTIONS, state="readonly")\
            .grid(row=0, column=0, sticky="ew")
        Badge(prio_row, self.prioridad, PRIORITY_COLORS).grid(row=0, column=1, padx=(10, 0), sticky="w")
        row(6, "Prioridad", prio_row)

        # Notebook: Días + Reporte
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, pady=(10, 0))
        self.nb = nb

        self.days_tab = ttk.Frame(nb, padding=10)
        nb.add(self.days_tab, text="Días / Data")

        days_box = ttk.LabelFrame(self.days_tab, text="Días", padding=10)
        days_box.pack(fill="both", expand=True)

        top_days = ttk.Frame(days_box)
        top_days.pack(fill="x", pady=(0, 8))
        ttk.Button(top_days, text="+ Agregar día", command=self.add_day).pack(side="left")

        self.notebook = ttk.Notebook(days_box)
        self.notebook.pack(fill="both", expand=True)
        self.day_tabs = []

        # Tab reporte (rich text)
        self.report_tab = ttk.Frame(nb, padding=10)
        nb.add(self.report_tab, text="Reporte (texto)")

        self.report_editor = ReportEditor(self.report_tab)
        self.report_editor.pack(fill="both", expand=True)

        actions = ttk.Frame(self)
        actions.pack(fill="x", pady=(12, 0))
        ttk.Button(actions, text="Guardar", command=self.on_save).pack(side="left")
        ttk.Button(actions, text="Exportar PDF", command=self.on_export).pack(side="left", padx=10)
        self.path_label = ttk.Label(actions, text="", foreground="#555")
        self.path_label.pack(side="right")

    def pick_logo(self):
        path = filedialog.askopenfilename(
            title="Elegir logo",
            filetypes=[("Imágenes", "*.png *.jpg *.jpeg *.gif *.webp"), ("Todos", "*.*")]
        )
        if path:
            self.logo_path.set(path)

    # ---------- días ----------
    def add_day(self):
        idx = len(self.day_tabs) + 1
        day_name = f"Día {idx}"
        self._create_day_tab(day_name, rows=[])

    def _create_day_tab(self, day_name: str, rows: list):
        frame = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(frame, text=day_name)

        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x", pady=(0, 6))

        ttk.Button(toolbar, text="+ Agregar fila", command=lambda: self.add_row(frame)).pack(side="left")
        ttk.Button(toolbar, text="Eliminar fila", command=lambda: self.delete_selected_row(frame)).pack(side="left", padx=8)
        ttk.Button(toolbar, text="Seleccionar carpeta (FILES)", command=lambda: self.set_folder_for_selected(frame)).pack(side="left", padx=8)
        ttk.Button(toolbar, text="Calcular peso", command=lambda: self.calc_weight_for_selected(frame)).pack(side="left", padx=8)

        cols = ("tarjeta", "dropbox", "backup_b", "backup_expo", "proxies", "files_path", "peso", "desglose", "comentarios")
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=12)

        headings = {
            "tarjeta": "TARJETA",
            "dropbox": "LINK DROPBOX",
            "backup_b": "BACKUP DISCO B",
            "backup_expo": "BACKUP DISCO EXPO",
            "proxies": "PROXIES",
            "files_path": "FILES (carpeta)",
            "peso": "PESO",
            "desglose": "DESGLOSE",
            "comentarios": "COMENTARIOS",
        }
        widths = {
            "tarjeta": 170,
            "dropbox": 240,
            "backup_b": 140,
            "backup_expo": 160,
            "proxies": 120,
            "files_path": 210,
            "peso": 90,
            "desglose": 240,
            "comentarios": 260,
        }

        for c in cols:
            tree.heading(c, text=headings[c])
            tree.column(c, width=widths[c], anchor="w")

        tree.pack(fill="both", expand=True)
        tree.bind("<Double-1>", lambda e: self.start_edit_cell(e, tree))

        tab = {"name": day_name, "frame": frame, "tree": tree, "rows": rows}
        self.day_tabs.append(tab)

        for row in rows:
            self._insert_tree_row(tree, row)

    def _insert_tree_row(self, tree, row: dict):
        vals = (
            row.get("tarjeta", ""),
            row.get("dropbox", ""),
            row.get("backup_b", "PENDIENTE"),
            row.get("backup_expo", "PENDIENTE"),
            row.get("proxies", "PENDIENTE"),
            row.get("files_path", ""),
            row.get("peso", ""),
            row.get("desglose", ""),
            row.get("comentarios", ""),
        )
        tree.insert("", "end", values=vals)

    def _tab_by_frame(self, frame):
        for t in self.day_tabs:
            if t["frame"] == frame:
                return t
        return None

    def _tab_by_tree(self, tree):
        for t in self.day_tabs:
            if t["tree"] == tree:
                return t
        return None

    def add_row(self, tab_frame):
        tab = self._tab_by_frame(tab_frame)
        if not tab:
            return
        row = {
            "tarjeta": "",
            "dropbox": "",
            "backup_b": "PENDIENTE",
            "backup_expo": "PENDIENTE",
            "proxies": "PENDIENTE",
            "files_path": "",
            "peso": "",
            "desglose": "",
            "comentarios": "",
        }
        tab["rows"].append(row)
        self._insert_tree_row(tab["tree"], row)

    def delete_selected_row(self, tab_frame):
        tab = self._tab_by_frame(tab_frame)
        if not tab:
            return
        tree = tab["tree"]
        sel = tree.selection()
        if not sel:
            return
        item = sel[0]
        idx = tree.index(item)
        tree.delete(item)
        if 0 <= idx < len(tab["rows"]):
            tab["rows"].pop(idx)

    def set_folder_for_selected(self, tab_frame):
        tab = self._tab_by_frame(tab_frame)
        if not tab:
            return
        tree = tab["tree"]
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Seleccionar fila", "Selecciona una fila primero.")
            return
        folder = filedialog.askdirectory(title="Seleccionar carpeta (FILES)")
        if not folder:
            return

        item = sel[0]
        idx = tree.index(item)
        values = list(tree.item(item, "values"))
        values[5] = folder
        tree.item(item, values=values)
        tab["rows"][idx]["files_path"] = folder

    def calc_weight_for_selected(self, tab_frame):
        tab = self._tab_by_frame(tab_frame)
        if not tab:
            return
        tree = tab["tree"]
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("Seleccionar fila", "Selecciona una fila primero.")
            return

        item = sel[0]
        idx = tree.index(item)
        row = tab["rows"][idx]
        folder = row.get("files_path", "")

        if not folder or not os.path.exists(folder):
            messagebox.showerror("Sin carpeta", "No hay carpeta válida en FILES para esa fila.")
            return

        size = folder_size_bytes(folder)
        h = human_size(size)

        values = list(tree.item(item, "values"))
        values[6] = h
        tree.item(item, values=values)
        row["peso"] = h

    def start_edit_cell(self, event, tree):
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        if not row_id or not col_id:
            return

        col_index = int(col_id.replace("#", "")) - 1
        columns = tree["columns"]
        col_name = columns[col_index]

        if col_name == "peso":
            return

        x, y, w, h = tree.bbox(row_id, col_id)
        value = tree.set(row_id, col_name)
        var = tk.StringVar(value=value)
        editor = None

        def save_and_close(_=None):
            new_val = var.get()
            tree.set(row_id, col_name, new_val)

            tab = self._tab_by_tree(tree)
            if tab:
                idx = tree.index(row_id)
                if 0 <= idx < len(tab["rows"]):
                    tab["rows"][idx][col_name] = new_val

            if editor:
                editor.destroy()

        if col_name in ("backup_b", "backup_expo", "proxies"):
            editor = ttk.Combobox(tree, textvariable=var, values=BACKUP_OPTIONS, state="readonly")
            editor.place(x=x, y=y, width=w, height=h)
            editor.focus_set()
            editor.bind("<<ComboboxSelected>>", save_and_close)
            editor.bind("<FocusOut>", save_and_close)
        else:
            editor = ttk.Entry(tree, textvariable=var)
            editor.place(x=x, y=y, width=w, height=h)
            editor.focus_set()
            editor.bind("<Return>", save_and_close)
            editor.bind("<FocusOut>", save_and_close)

    # ---------- portada ----------
    def _ui_dates_to_iso(self, fi_ui: str, ff_ui: str):
        fi_ui = fi_ui.strip()
        ff_ui = ff_ui.strip()
        if fi_ui and not valid_date_mmddyyyy(fi_ui):
            raise ValueError("Fecha de inicio debe ser MM/DD/YYYY (ej: 01/13/2026).")
        if ff_ui and not valid_date_mmddyyyy(ff_ui):
            raise ValueError("Fecha de fin debe ser MM/DD/YYYY (ej: 01/13/2026).")
        fi_iso = mmddyyyy_to_iso(fi_ui) if fi_ui else ""
        ff_iso = mmddyyyy_to_iso(ff_ui) if ff_ui else ""
        return fi_iso, ff_iso

    def collect_data(self) -> dict:
        fi_ui = self.fecha_inicio_ui.get()
        ff_ui = self.fecha_fin_ui.get()
        fi_iso, ff_iso = self._ui_dates_to_iso(fi_ui, ff_ui)

        dias = []
        for t in self.day_tabs:
            dias.append({"nombre": t["name"], "rows": t["rows"]})

        return {
            "logo_path": self.logo_path.get().strip(),
            "status_general": self.status_general.get().strip(),
            "proyecto": self.proyecto.get().strip(),
            "fecha_inicio": fi_iso,
            "fecha_fin": ff_iso,
            "encargado": self.encargado.get().strip(),
            "etapa": self.etapa.get().strip(),
            "prioridad": self.prioridad.get().strip(),
            "dias": dias,
            "reporte_rich": self.report_editor.get_payload(),
        }

    def load_from_data(self, data: dict):
        self.logo_path.set(data.get("logo_path", ""))
        self.status_general.set(data.get("status_general", STATUS_OPTIONS[0]))
        self.proyecto.set(data.get("proyecto", ""))
        self.encargado.set(data.get("encargado", ""))
        self.etapa.set(data.get("etapa", ""))
        self.prioridad.set(data.get("prioridad", PRIORITY_OPTIONS[1]))

        fi_iso = (data.get("fecha_inicio") or "").strip()
        ff_iso = (data.get("fecha_fin") or "").strip()

        def to_ui(iso_s: str) -> str:
            if not iso_s:
                return ""
            try:
                return iso_to_mmddyyyy(iso_s)
            except Exception:
                return iso_s

        self.fecha_inicio_ui.set(to_ui(fi_iso))
        self.fecha_fin_ui.set(to_ui(ff_iso))

        # reset days tabs
        for tab_id in list(self.notebook.tabs()):
            self.notebook.forget(tab_id)
        self.day_tabs = []

        dias = data.get("dias", [])
        if not dias:
            dias = [{"nombre": "Día 1", "rows": []}]
        for d in dias:
            self._create_day_tab(d.get("nombre", "Día"), d.get("rows", []))

        self.report_editor.set_payload(data.get("reporte_rich", {"text": "", "bold_ranges": [], "italic_ranges": []}))

        p = self.controller.current_path
        self.path_label.config(text=(p if p else "Sin guardar"))

    def on_save(self):
        try:
            data = self.collect_data()
        except ValueError as ve:
            messagebox.showerror("Fecha inválida", str(ve))
            return
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        path = self.controller.save_report(data)
        if path:
            self.controller.report_data = data
            self.path_label.config(text=path)
            messagebox.showinfo("Guardado", f"Reporte guardado.\n\nUbicación:\n{path}")

    def on_export(self):
        try:
            data = self.collect_data()
        except ValueError as ve:
            messagebox.showerror("Fecha inválida", str(ve))
            return
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.controller.report_data = data
        self.controller.export_pdf(data)


# =========================
#  MAIN
# =========================
if __name__ == "__main__":
    ensure_dirs()
    app = ReportApp()
    app.mainloop()
