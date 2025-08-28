# neeeew.py  — BALADA PACKAGING · Presupuestos (Python 3.10, Windows, PyQt6)
# ---------------------------------------------------------------------------------
# Cambios:
# - Tablas HTML en correo / preview / PDF.
# - Firma con Nombre/Empresa desde Configuración.
# - Panel de sistema (CPU/RAM/GPU/CUDA) formateado.
# - Botón de actualizar grande.
# - Campos de rutas (TGP y PDF) con "Elegir carpeta de destino".
# - Logs detallados y correcciones previas mantenidas.
# ---------------------------------------------------------------------------------

import os
import re
import sys
import json
import time
import smtplib
import shutil
import platform
import subprocess
from typing import List, Dict, Any, Optional, Tuple

# Qt
from PyQt6.QtCore import (
    Qt, QSize, QTimer, QThread, pyqtSignal, QSettings, QMarginsF
)
from PyQt6.QtGui import (
    QIcon, QPixmap, QImage, QFont, QTextDocument
)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QStackedWidget, QLineEdit, QTextEdit,
    QComboBox, QMessageBox, QGroupBox, QFormLayout, QTabWidget, QDialog
)
from PyQt6.QtPrintSupport import QPrinter

# IA y PDF
import fitz  # PyMuPDF
import google.generativeai as genai
from google.generativeai.types import GenerationConfig

# HTTP/updates
import requests

# Cámara
import cv2

# Opcional para info del sistema (RAM, GPU)
try:
    import psutil
except Exception:
    psutil = None

# ------------------- Utilidades básicas -------------------

APP_ORG = "BitStation"
APP_NAME = "Presupuestos"
GITHUB_USER = "BitStationBusiness"
GITHUB_REPO = "BitStation_Budget_Quotation_version_of_BALADA_PACKAGING"

def get_current_version() -> str:
    vfile = "version.txt"
    if not os.path.exists(vfile):
        with open(vfile, "w", encoding="utf-8") as f:
            f.write("0.1")
        return "0.1"
    try:
        with open(vfile, "r", encoding="utf-8") as f:
            v = f.read().strip() or "0.1"
        return v
    except Exception:
        return "0.1"

APP_VERSION = get_current_version()

def resource_path(*parts) -> str:
    return os.path.join(os.getcwd(), *parts)

def human_ex(err: Exception) -> str:
    return f"{type(err).__name__}: {err}"

# ------------------- Update / Version ---------------------

class UpdateCheckerWorker(QThread):
    finished = pyqtSignal(dict)
    def run(self):
        api_url = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest"
        result = {"ok": False, "error": "", "tag": "", "assets": []}
        try:
            r = requests.get(api_url, timeout=10)
            r.raise_for_status()
            data = r.json()
            tag = (data.get("tag_name") or "").lstrip("v")
            assets = data.get("assets", [])
            result.update({"ok": True, "tag": tag, "assets": assets})
        except Exception as e:
            result["error"] = human_ex(e)
        self.finished.emit(result)

# ------------------- Widgets reusables --------------------

class ClickLabel(QLabel):
    clicked = pyqtSignal()
    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()

class DropImage(QLabel):
    image_loaded = pyqtSignal(str)  # path
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet("border:2px dashed #b7bcc7; border-radius:10px; background:#fafafa; color:#999;")
        self.setMinimumSize(480, 420)
        self._placeholder()

        self._pix: Optional[QPixmap] = None
        self.bytes: Optional[bytes] = None
        self.mime: Optional[str] = None
        self.loaded_path: Optional[str] = None

    # Placeholder helpers
    def _placeholder(self):
        self.setPixmap(QPixmap())
        self.setText("Arrastra una imagen/PDF aquí")

    def reset_placeholder(self):
        self._pix = None
        self.bytes = None
        self.mime = None
        self.loaded_path = None
        self._placeholder()

    # DnD
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            e.ignore()

    def dropEvent(self, e):
        files = [u.toLocalFile() for u in e.mimeData().urls()]
        if files:
            self.load_path(files[0])

    def resizeEvent(self, ev):
        super().resizeEvent(ev)
        self._rescale()

    def _rescale(self):
        if self._pix:
            self.setPixmap(self._pix.scaled(self.size(), Qt.AspectRatioMode.KeepAspectRatio,
                                            Qt.TransformationMode.SmoothTransformation))

    def load_path(self, path: str):
        try:
            self.loaded_path = path
            if path.lower().endswith(".pdf"):
                with fitz.open(path) as doc:
                    page = doc.load_page(0)
                    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
                    png = pix.tobytes("png")
                    self.bytes = png
                    self.mime = "image/png"
                    qimg = QImage.fromData(png)
            else:
                with open(path, "rb") as f:
                    self.bytes = f.read()
                ext = os.path.splitext(path)[1].lower()
                self.mime = "image/jpeg" if ext in (".jpg", ".jpeg") else "image/bmp" if ext == ".bmp" else "image/png"
                qimg = QImage(path)

            if qimg.isNull():
                raise ValueError("No se pudo cargar la previsualización")
            self._pix = QPixmap.fromImage(qimg)
            self.setText("")  # quita placeholder
            self._rescale()
            print(f"[DropImage] Imagen cargada: {path}")
            self.image_loaded.emit(path)
        except Exception as e:
            self._pix = None
            self.bytes = None
            self.mime = None
            self._placeholder()
            print("[DropImage] Error:", human_ex(e))

# ------------------- IA helpers (comunes) --------------

import unicodedata
def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return s.upper()

NUM_PAT = re.compile(r"(?:\d{1,3}(?:[.,]\d{3})+|\d+)(?:[.,]\d{2,3})?")
def first_number_or_blank(x: Any) -> str:
    if x is None: return ""
    m = NUM_PAT.search(str(x))
    return m.group(0) if m else ""

def has_digit(s: str) -> bool:
    return any(ch.isdigit() for ch in str(s))

def safe_json_loads(s: str) -> Dict[str, Any]:
    s = s.strip()
    m = re.search(r"```(?:json)?\s*\n([\s\S]+?)\n```", s, flags=re.I)
    if m: s = m.group(1).strip()
    try:
        return json.loads(s or "{}")
    except Exception:
        # Algunos modelos devuelven comillas simples: intenta arreglar rápido
        try:
            return json.loads(s.replace("'", '"'))
        except Exception:
            return {}

def html_escape(t: str) -> str:
    return (t or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

# Conversión rápida de tablas Markdown a HTML (fallback)
def markdown_table_to_html(text: str) -> Optional[str]:
    lines = [ln.rstrip() for ln in (text or "").splitlines()]
    rows = [ln for ln in lines if ln.strip().startswith("|") and ln.strip().endswith("|")]
    if not rows:
        return None
    # Elimina la línea separadora si existe (----)
    if len(rows) >= 2 and set(rows[1].replace("|","").strip()) <= set("-: "):
        rows.pop(1)
    def split_row(r): return [c.strip() for c in r.strip("|").split("|")]
    cells = [split_row(r) for r in rows]
    if not cells: return None
    header = cells[0]
    body = cells[1:] if len(cells) > 1 else []
    # Construye HTML simple compatible con correo
    def td(tag, v): return f"<{tag} style='border:1px solid #ccc;padding:6px 8px'>{html_escape(v)}</{tag}>"
    thead = "<tr>" + "".join(td("th", h) for h in header) + "</tr>"
    tbody = "".join("<tr>"+ "".join(td("td", v) for v in row) + "</tr>" for row in body)
    table = (
        "<table style='border-collapse:collapse;border:1px solid #ccc;font-family:Arial;font-size:12px'>"
        f"<thead>{thead}</thead><tbody>{tbody}</tbody></table>"
    )
    # Devuelve el bloque completo reemplazando la sección en el texto original
    html_blocks = []
    in_table = False
    for ln in lines:
        if ln.strip().startswith("|") and ln.strip().endswith("|"):
            if not in_table:
                in_table = True
                continue
            else:
                continue
        else:
            if in_table:
                in_table = False
                html_blocks.append(table)
            html_blocks.append(html_escape(ln))
    if in_table:
        html_blocks.append(table)
    return "<br>".join(html_blocks)

# ------------------- Opción A (Extractor TGP) ----------

class OptionAPage(QWidget):
    """Extractor de datos de facturas tipo TGP (PDF 1:1 / Imagen con IA)."""
    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        self.api_key: Optional[str] = None

        self.bytes: Optional[bytes] = None
        self.mime: Optional[str] = None
        self.src_path: Optional[str] = None
        self.pdf_rows: Optional[List[Dict[str, str]]] = None

        self.prompt_count_a = (
            "Cuenta cuántos RENGLONES visibles conforman el CUERPO de la tabla de productos de la factura. "
            "Cada renglón impreso = una fila. EXCLUYE encabezados (DESCRIPCIÓN/CANTIDAD/PRECIO) "
            "y EXCLUYE totales (BRUTO, BASE, TOTAL...). "
            'Devuelve SOLO JSON: {"row_count": <entero>}'
        ).strip()
        self.prompt_count_b = (
            "Identifica TODAS las LÍNEAS impresas del CUERPO de la tabla de productos (una por renglón, "
            "sin encabezados ni totales). Para cada línea, devuelve su y_center aproximado (0.0–1.0) y el texto."
            ' Salida SOLO JSON: {"lines":[{"y":0.0,"text":"..."}]}'
        ).strip()
        self.prompt_rows_image_strict = (
            "Devuelve EXACTAMENTE N renglones de la TABLA DE PRODUCTOS, en el MISMO orden visual.\n"
            "- N = <<N_ROWS>> (no devuelvas más ni menos).\n"
            "- REGLA CLAVE: cada renglón impreso en la imagen = UNA fila en el JSON. No fusiones líneas contiguas.\n"
            "- Define 3 bandas de columna desde la cabecera (DESCRIPCIÓN · CANTIDAD · PRECIO) y asigna tokens por posición.\n"
            "- Cualquier token no numérico en CANTIDAD/PRECIO muévelo a 'descripcion'.\n"
            "- 'cantidad' y 'precio' deben contener SOLO números con coma/punto o ''.\n"
            "- EXCLUYE encabezados y EXCLUYE totales.\n"
            'Salida SOLO JSON: {"rows":[{"descripcion":"string","cantidad":"string","precio":"string"}]}'
        ).strip()

        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)

        # Topbar
        tb = QHBoxLayout()
        self.btn_back = QPushButton("←"); self.btn_back.setFixedSize(36, 28)
        self.btn_back.clicked.connect(self.main.go_home)
        title = QLabel("BALADA PACKAGING .S.L.U.")
        f = QFont(); f.setPointSize(16); f.setBold(True); title.setFont(f)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        gear = QPushButton("⚙️"); gear.setFixedSize(36, 28)
        gear.clicked.connect(self.main.open_settings)
        tb.addWidget(self.btn_back)
        tb.addStretch(1); tb.addWidget(title); tb.addStretch(1)
        tb.addWidget(gear)
        root.addLayout(tb)

        # Contenido
        content = QHBoxLayout()
        self.drop = DropImage(self)
        self.drop.image_loaded.connect(self._on_loaded)

        # Tabla
        from PyQt6.QtWidgets import QTableWidget, QTableWidgetItem, QHeaderView
        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Descripción", "Cantidad", "Precio"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)

        content.addWidget(self.drop, 1)
        content.addWidget(self.table, 1)
        root.addLayout(content)

        # Bottom
        bb = QHBoxLayout()
        self.btn_open = QPushButton("Abrir archivo…")
        self.btn_open.clicked.connect(self._open_file)
        self.btn_proc = QPushButton("Procesar")
        self.btn_proc.clicked.connect(self._process)
        self.btn_proc.setEnabled(False)
        bb.addStretch(1); bb.addWidget(self.btn_open); bb.addWidget(self.btn_proc)
        root.addLayout(bb)

        # Timer para tomar API de settings
        self._api_timer = QTimer(self); self._api_timer.setInterval(800); self._api_timer.setSingleShot(True)
        self._api_timer.timeout.connect(self._refresh_api_key)
        self._api_timer.start()

    def _refresh_api_key(self):
        key = self.main.settings.value("accounts/google_api_key", "", str).strip()
        self.api_key = key or None
        print(f"[TGP] API Key cargada: {'sí' if self.api_key else 'no'}")
        self.btn_proc.setEnabled(bool(self.api_key and self.drop.bytes))

    def _on_loaded(self, path: str):
        self.src_path = path
        self.pdf_rows = None
        if path.lower().endswith(".pdf"):
            try:
                self.pdf_rows = self._rows_from_pdf(path)
            except Exception as e:
                print("[TGP][PDF] Parser error:", human_ex(e))
        self._refresh_api_key()
        if self.api_key:
            self._process()

    def _open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Selecciona imagen o PDF", "",
            "Archivos soportados (*.pdf *.png *.jpg *.jpeg *.bmp);;Todos (*.*)"
        )
        if path:
            self.drop.load_path(path)

    def _process(self):
        if not (self.api_key and self.drop.bytes and self.drop.mime):
            return
        QApplication.setOverrideCursor(Qt.CursorShape.BusyCursor)
        try:
            if self.src_path and self.src_path.lower().endswith(".pdf") and self.pdf_rows:
                rows = self.pdf_rows
            else:
                n = self._count_rows_dual()
                rows = self._extract_rows_image(n)
                if len(rows) != n:
                    rows = self._repair_to_n(rows, n)
            self._populate(rows)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error en el proceso:\n{human_ex(e)}")
        finally:
            QApplication.restoreOverrideCursor()

    def _populate(self, rows: List[Dict[str, Any]]):
        from PyQt6.QtWidgets import QTableWidgetItem
        self.table.setRowCount(0)
        for obj in rows:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(str(obj.get("descripcion", ""))))
            self.table.setItem(r, 1, QTableWidgetItem(str(obj.get("cantidad", ""))))
            self.table.setItem(r, 2, QTableWidgetItem(str(obj.get("precio", ""))))

    # ---- PDF parser 1:1
    def _rows_from_pdf(self, pdf_path: str) -> Optional[List[Dict[str, str]]]:
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)

        words = page.get_text("words")
        if not words:
            return None

        lines: Dict[int, List[Tuple]] = {}
        for x0, y0, x1, y1, txt, bno, lno, wno in words:
            key = int(round(y0 / 2.0))
            lines.setdefault(key, []).append((x0, y0, x1, y1, txt))

        header_key = None
        desc_c = qty_c = price_c = codigo_c = None
        for k, ws in sorted(lines.items()):
            texts = [_norm(t[4]) for t in ws]
            if any("DESCRIPCION" in t for t in texts) and any("CANTIDAD" in t for t in texts) and any("PRECIO" in t for t in texts):
                for x0, y0, x1, y1, txt in ws:
                    n = _norm(txt)
                    mid = (x0 + x1) / 2
                    if "CODIGO" in n and codigo_c is None: codigo_c = mid
                    if "DESCRIPCION" in n and desc_c is None: desc_c = mid
                    if "CANTIDAD" in n and qty_c is None: qty_c = mid
                    if "PRECIO" in n and price_c is None: price_c = mid
                header_key = k
                break
        if header_key is None or not (desc_c and qty_c and price_c):
            return None

        header_y = max(w[3] for w in lines[header_key])
        bottom = page.rect.height
        for k, ws in sorted(lines.items()):
            if k <= header_key:
                continue
            t = " ".join(_norm(w[4]) for w in ws)
            if any(word in t for word in ("BRUTO", "TOTAL", "BASE")):
                bottom = min(bottom, min(w[1] for w in ws))
                break

        split_cd = (codigo_c + desc_c) / 2.0 if codigo_c else (desc_c - (qty_c - desc_c) * 0.9)
        split_dq = (desc_c + qty_c) / 2.0
        split_qp = (qty_c + price_c) / 2.0
        right_guard = price_c + (price_c - qty_c)

        rows: List[Dict[str, str]] = []
        for k, ws in sorted(lines.items()):
            y_top = min(w[1] for w in ws)
            y_bot = max(w[3] for w in ws)
            if y_top <= header_y or y_bot >= bottom:
                continue

            selected = []
            for x0, y0, x1, y1, txt in ws:
                mid = (x0 + x1) / 2.0
                if (split_cd - 20) <= mid <= right_guard:
                    selected.append((x0, y0, x1, y1, txt, mid))
            if not selected:
                continue
            selected.sort(key=lambda t: t[0])

            desc_parts, qty_parts, price_parts = [], [], []
            for *_x0, _y0, _x1, _y1, txt, mid in selected:
                if mid < split_dq:
                    desc_parts.append(txt)
                elif mid < split_qp:
                    qty_parts.append(txt)
                else:
                    price_parts.append(txt)

            qty_non = [t for t in qty_parts if not has_digit(t)]
            price_non = [t for t in price_parts if not has_digit(t)]
            if qty_non:
                desc_parts.extend(qty_non)
                qty_parts = [t for t in qty_parts if has_digit(t)]
            if price_non:
                desc_parts.extend(price_non)
                price_parts = [t for t in price_parts if has_digit(t)]

            desc = " ".join(desc_parts).strip()
            qty = first_number_or_blank(" ".join(qty_parts).strip())
            price = first_number_or_blank(" ".join(price_parts).strip())
            if not desc:
                continue
            rows.append({"descripcion": desc, "cantidad": qty, "precio": price})
        return rows or None

    # ---- IA dual conteo + extracción
    def _count_rows_dual(self) -> int:
        genai.configure(api_key=self.api_key)
        model = genai.GenerativeModel("gemini-1.5-flash-latest")
        part = {"mime_type": self.drop.mime, "data": self.drop.bytes}

        # Conteo A
        schema_a = {"type": "OBJECT","properties": {"row_count": {"type": "INTEGER"}},"required": ["row_count"]}
        cfg_a = GenerationConfig(response_mime_type="application/json", response_schema=schema_a, temperature=0.0)
        resp_a = model.generate_content([self.prompt_count_a, part], generation_config=cfg_a)
        raw_a = getattr(resp_a, "text", "") or "{}"
        try: n_a = int(safe_json_loads(raw_a).get("row_count", 0))
        except Exception: n_a = 0

        # Conteo B
        schema_b = {
            "type": "OBJECT",
            "properties": {"lines": {"type": "ARRAY","items": {"type": "OBJECT","properties": {"y": {"type": "NUMBER"}, "text": {"type": "STRING"}}, "required": ["y","text"]}}},
            "required": ["lines"]
        }
        cfg_b = GenerationConfig(response_mime_type="application/json", response_schema=schema_b, temperature=0.0)
        resp_b = model.generate_content([self.prompt_count_b, part], generation_config=cfg_b)
        raw_b = getattr(resp_b, "text", "") or "{}"
        try:
            lines = safe_json_loads(raw_b).get("lines", [])
            lines = [it for it in lines if str(it.get("text", "")).strip()]
            lines.sort(key=lambda z: z.get("y", 0))
            dedup, last_y = [], None
            for it in lines:
                y = it.get("y", 0.0)
                if last_y is None or abs(y - last_y) > 0.01:
                    dedup.append(it); last_y = y
            n_b = len(dedup)
        except Exception:
            n_b = 0

        n = max(n_a, n_b)
        return n if n > 0 else 1

    def _extract_rows_image(self, n: int) -> List[Dict[str, str]]:
        genai.configure(api_key=self.api_key)
        model = genai.GenerativeModel("gemini-1.5-flash-latest")
        schema = {
            "type": "OBJECT",
            "properties": {
                "rows": {"type": "ARRAY","items": {"type": "OBJECT","properties": {"descripcion": {"type": "STRING"}, "cantidad": {"type": "STRING"}, "precio": {"type": "STRING"}}, "required": ["descripcion","cantidad","precio"]}}
            },
            "required": ["rows"]
        }
        cfg = GenerationConfig(response_mime_type="application/json", response_schema=schema, temperature=0.0)

        prompt = self.prompt_rows_image_strict.replace("<<N_ROWS>>", str(n))
        part = {"mime_type": self.drop.mime, "data": self.drop.bytes}
        resp = model.generate_content([prompt, part], generation_config=cfg)
        raw = getattr(resp, "text", "") or "{}"
        data = safe_json_loads(raw)
        rows = data.get("rows", [])
        clean = [{"descripcion": str(r.get("descripcion", "")), "cantidad": first_number_or_blank(r.get("cantidad", "")), "precio": first_number_or_blank(r.get("precio", ""))} for r in rows]
        return clean

    def _repair_to_n(self, rows: List[Dict[str, str]], n: int) -> List[Dict[str, str]]:
        genai.configure(api_key=self.api_key)
        model = genai.GenerativeModel("gemini-1.5-flash-latest")
        schema = {
            "type": "OBJECT",
            "properties": {"rows": {"type": "ARRAY","items": {"type": "OBJECT","properties": {"descripcion": {"type": "STRING"}, "cantidad": {"type": "STRING"}, "precio": {"type": "STRING"}}, "required": ["descripcion","cantidad","precio"]}}},
            "required": ["rows"]
        }
        cfg = GenerationConfig(response_mime_type="application/json", response_schema=schema, temperature=0.0)

        prompt = (
            "Corrige el siguiente JSON para que tenga EXACTAMENTE N filas, sin inventar texto, "
            "separando cualquier fila que haya fusionado dos renglones impresos. Mantén el orden visual. "
            f"N = {n}\nJSON actual:\n{json.dumps({'rows': rows}, ensure_ascii=False)}\n"
            'Salida SOLO JSON con {"rows":[...]}.'
        )
        part = {"mime_type": self.drop.mime, "data": self.drop.bytes}
        resp = model.generate_content([prompt, part], generation_config=cfg)
        raw = getattr(resp, "text", "") or "{}"
        data = safe_json_loads(raw)
        fixed = [{"descripcion": str(r.get("descripcion", "")), "cantidad": first_number_or_blank(r.get("cantidad", "")), "precio": first_number_or_blank(r.get("precio", ""))} for r in data.get("rows", [])]
        if len(fixed) < n:
            fixed += [{"descripcion": "", "cantidad": "", "precio": ""} for _ in range(n - len(fixed))]
        if len(fixed) > n:
            fixed = fixed[:n]
        return fixed

# ------------------- IA worker (Opción B) ------------

class EmailGenWorker(QThread):
    finished = pyqtSignal(str, str, str)  # subject, html, text
    failed = pyqtSignal(str)
    def __init__(self, api_key: str, recipient_name: str, sender_name: str, img_bytes: bytes, mime: str):
        super().__init__()
        self.api_key = api_key
        self.recipient_name = recipient_name.strip()
        self.sender_name = sender_name.strip()
        self.img_bytes = img_bytes
        self.mime = mime or "image/png"
    def run(self):
        try:
            genai.configure(api_key=self.api_key)
            model = genai.GenerativeModel("gemini-1.5-flash-latest")
            schema = {
                "type": "OBJECT",
                "properties": {
                    "subject": {"type": "STRING"},
                    "body_html": {"type": "STRING"},
                    "body_text": {"type": "STRING"}
                },
                "required": ["subject","body_html"]
            }
            cfg = GenerationConfig(response_mime_type="application/json", response_schema=schema, temperature=0.2)
            prompt = (
                "Lee el texto manuscrito/impreso de la imagen y genera un correo de SOLICITUD DE COTIZACIÓN.\n"
                f"- Saluda al destinatario por su nombre si está disponible: {self.recipient_name or '[Nombre del destinatario]'}.\n"
                "- Devuelve SIEMPRE una versión HTML (body_html) y, si puedes, también texto plano (body_text).\n"
                "- IMPORTANTÍSIMO: en HTML usa una TABLA <table> con borde (border-collapse), estilo inline, 3 columnas: Ítem | Cantidad | Detalle.\n"
                "- No inventes datos; si faltan, usa descripciones genéricas.\n"
                "- Termina con 'Atentamente,' y un marcador [Tu Nombre] que luego yo reemplazo por el nombre/empresa del remitente.\n"
                "Devuelve SOLO JSON con 'subject', 'body_html' y opcionalmente 'body_text'."
            )
            part = {"mime_type": self.mime, "data": self.img_bytes}
            resp = model.generate_content([prompt, part], generation_config=cfg)
            raw = getattr(resp, "text", "") or "{}"
            data = safe_json_loads(raw)
            subject = (data.get("subject") or "").strip() or "Solicitud de cotización"
            body_html = (data.get("body_html") or "").strip()
            body_text = (data.get("body_text") or "").strip()

            # Reemplazo de destinatario y remitente
            if self.recipient_name:
                body_html = body_html.replace("[Nombre del destinatario]", self.recipient_name)
                body_text = body_text.replace("[Nombre del destinatario]", self.recipient_name)
            if self.sender_name:
                body_html = body_html.replace("[Tu Nombre]", self.sender_name)
                body_text = body_text.replace("[Tu Nombre]", self.sender_name)

            # Fallback: si no vino HTML, intenta convertir tablas estilo Markdown
            if not body_html:
                pt_html = html_escape(body_text).replace("\n", "<br>")
                body_html = markdown_table_to_html(body_text) or f"<div style='font-family:Arial'>{pt_html}</div>"

            if not body_text:
                # Texto simple a partir del HTML (súper básico)
                body_text = re.sub("<[^<]+?>", "", body_html).replace("&nbsp;"," ").replace("&amp;","&")

            self.finished.emit(subject, body_html, body_text)
        except Exception as e:
            self.failed.emit(human_ex(e))

# ------------------- Opción B (Presupuestos) ---------

class OptionBPage(QWidget):
    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        self._ai_worker: Optional[EmailGenWorker] = None
        self._last_html: str = ""
        self._last_text: str = ""
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)

        # Topbar
        tb = QHBoxLayout()
        self.btn_back = QPushButton("←"); self.btn_back.setFixedSize(36, 28)
        self.btn_back.clicked.connect(self.main.go_home)
        title = QLabel("BALADA PACKAGING .S.L.U.")
        f = QFont(); f.setPointSize(16); f.setBold(True); title.setFont(f)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        gear = QPushButton("⚙️"); gear.setFixedSize(36, 28)
        gear.clicked.connect(self.main.open_settings)
        tb.addWidget(self.btn_back)
        tb.addStretch(1); tb.addWidget(title); tb.addStretch(1)
        tb.addWidget(gear)
        root.addLayout(tb)

        # Layout columnas
        cols = QHBoxLayout()
        root.addLayout(cols, 1)

        # Izquierda: drop + cámara
        left = QVBoxLayout()
        self.drop = DropImage(self)
        left.addWidget(self.drop, 1)

        cam_row = QHBoxLayout()
        self.cam_combo = QComboBox()
        self.cam_combo.addItem("Ninguna (arrastrar imagen)", -1)
        self.btn_capture = QPushButton("Capturar imagen")
        self.btn_capture.clicked.connect(self._capture_frame)
        self._populate_cameras()
        self.cam_combo.currentIndexChanged.connect(self._on_cam_changed)

        cam_row.addWidget(self.cam_combo, 2)
        cam_row.addWidget(self.btn_capture, 1)
        left.addLayout(cam_row)

        cols.addLayout(left, 1)

        # Derecha: formulario
        right = QVBoxLayout()

        def make_line(tt: str):
            lbl = QLabel(tt); lbl.setStyleSheet("color:#666;")
            le = QLineEdit()
            return lbl, le

        lbl_to, self.to_edit = make_line("PERSONA O NOMBRE DE EMPRESA A QUIEN VA DIRIGIDO")
        lbl_mail, self.mail_edit = make_line("CORREO DE CONTACTO")
        lbl_subj, self.subject_edit = make_line("ASUNTO DE CORREO")

        right.addWidget(lbl_to); right.addWidget(self.to_edit)
        right.addWidget(lbl_mail); right.addWidget(self.mail_edit)
        right.addWidget(lbl_subj); right.addWidget(self.subject_edit)

        self.body_edit = QTextEdit()
        self.body_edit.setPlaceholderText("Se generará un cuerpo en HTML con tabla al cargar una imagen.")
        right.addWidget(self.body_edit, 1)

        actions = QHBoxLayout()
        self.btn_pdf = QPushButton("DESCARGAR PDF")
        self.btn_send = QPushButton("ENVIAR")
        self.btn_pdf.clicked.connect(self.export_pdf)
        self.btn_send.clicked.connect(self.send_email)
        actions.addStretch(1); actions.addWidget(self.btn_pdf); actions.addWidget(self.btn_send)
        right.addLayout(actions)

        cols.addLayout(right, 1)

        # Auto-regeneración al cambiar nombre (debounce)
        self._regen_timer = QTimer(self); self._regen_timer.setSingleShot(True)
        self._regen_timer.timeout.connect(lambda: self._maybe_autogenerate(trigger="name-change"))
        self.to_edit.textChanged.connect(lambda: self._regen_timer.start(500))

        # Regenerar cuando se cargue imagen
        self.drop.image_loaded.connect(lambda _p: self._maybe_autogenerate(trigger="image"))

    # Cámaras
    def _populate_cameras(self):
        while self.cam_combo.count() > 1:
            self.cam_combo.removeItem(1)

        found = 0
        backends = [cv2.CAP_DSHOW, cv2.CAP_MSMF]
        for idx in range(0, 6):
            ok_idx = False
            for be in backends:
                cap = cv2.VideoCapture(idx, be)
                if cap.isOpened():
                    ret, _ = cap.read()
                    cap.release()
                    if ret:
                        self.cam_combo.addItem(f"Cámara {idx}", idx)
                        print(f"[Cam] Detectada cámara #{idx}")
                        found += 1
                        ok_idx = True
                        break
                else:
                    cap.release()
            if not ok_idx:
                continue

        if hasattr(self, "btn_capture"):
            self.btn_capture.setEnabled(found > 0)
        print(f"[Cam] Total cámaras disponibles: {found}")

    def _on_cam_changed(self, idx: int):
        cam_index = self.cam_combo.currentData()
        print(f"[Cam] Cambio de selección: {cam_index}")
        if cam_index is None or int(cam_index) == -1:
            self.drop.reset_placeholder()

    def _capture_frame(self):
        cam_index = self.cam_combo.currentData()
        if cam_index is None or cam_index == -1:
            QMessageBox.warning(self, "Cámara", "Selecciona una cámara distinta de 'Ninguna'.")
            return
        cap = cv2.VideoCapture(int(cam_index), cv2.CAP_DSHOW)
        if not cap.isOpened():
            QMessageBox.critical(self, "Cámara", "No se pudo abrir la cámara seleccionada.")
            return
        ret, frame = cap.read()
        cap.release()
        if not ret:
            QMessageBox.critical(self, "Cámara", "No se pudo capturar imagen.")
            return
        tmp = resource_path("_captura_tmp.png")
        cv2.imwrite(tmp, frame)
        self.drop.load_path(tmp)
        try: os.remove(tmp)
        except Exception: pass

    # IA
    def _img_and_mime_or_none(self):
        return (self.drop.bytes, self.drop.mime) if self.drop.bytes else (None, None)

    def _maybe_autogenerate(self, trigger: str):
        api_key = self.main.settings.value("accounts/google_api_key", "", str)
        company = self.main.settings.value("accounts/company_name", "", str)
        print(f"[IA-Presupuestos] Trigger: {trigger} · API {'sí' if api_key else 'no'}")
        if not api_key:
            return

        # Si sólo cambia el nombre del destinatario y ya hay HTML, reemplaza inline
        if trigger.startswith("name") and self._last_html:
            name_now = self.to_edit.text().strip()
            html = self._last_html
            if name_now:
                html = html.replace("[Nombre del destinatario]", name_now)
            if company:
                html = html.replace("[Tu Nombre]", company)
            self._last_html = html
            self.body_edit.blockSignals(True); self.body_edit.setHtml(html); self.body_edit.blockSignals(False)
            return

        img, mime = self._img_and_mime_or_none()
        if not img:
            print("[IA-Presupuestos] No hay imagen para procesar.")
            return

        name_now = self.to_edit.text().strip()
        self.subject_edit.setPlaceholderText("Generando asunto con IA…")
        self.body_edit.setPlaceholderText("Generando cuerpo (HTML) con IA…")
        self.btn_pdf.setEnabled(False); self.btn_send.setEnabled(False)

        self._ai_worker = EmailGenWorker(api_key, name_now, company, img, mime)
        self._ai_worker.finished.connect(self._on_ai_ready)
        self._ai_worker.failed.connect(self._on_ai_failed)
        self._ai_worker.start()
        print("[IA-Presupuestos] Worker iniciado.")

    def _on_ai_ready(self, subject: str, body_html: str, body_text: str):
        print("[IA-Presupuestos] Resultado recibido.")
        self._last_html, self._last_text = body_html, body_text
        self.subject_edit.blockSignals(True); self.body_edit.blockSignals(True)
        if not self.subject_edit.text().strip():
            self.subject_edit.setText(subject)
        self.body_edit.setHtml(body_html)
        self.subject_edit.blockSignals(False); self.body_edit.blockSignals(False)
        self.btn_pdf.setEnabled(True); self.btn_send.setEnabled(True)

    def _on_ai_failed(self, msg: str):
        print("[IA-Presupuestos] Error:", msg)
        self.btn_pdf.setEnabled(True); self.btn_send.setEnabled(True)

    # Helpers de cuerpo actual
    def _current_body_html_text(self) -> Tuple[str, str]:
        html = self._last_html or self.body_edit.toHtml()
        txt = self._last_text or re.sub("<[^<]+?>", "", html).replace("&nbsp;"," ").replace("&amp;","&")
        # Asegura firma con nombre/empresa de configuración
        company = self.main.settings.value("accounts/company_name", "", str)
        if company:
            html = html.replace("[Tu Nombre]", company)
            txt = txt.replace("[Tu Nombre]", company)
        return html, txt

    # PDF y correo
    def export_pdf(self):
        default_dir = self.main.settings.value("paths/pdf_dir", os.getcwd(), str) or os.getcwd()
        default_path = os.path.join(default_dir, "presupuesto.pdf")
        path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF", default_path, "PDF (*.pdf)")
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(path)
            printer.setPageMargins(QMarginsF(12, 12, 12, 12))

            subj = (self.subject_edit.text().strip() or "Presupuesto")
            body_html, _ = self._current_body_html_text()

            html = f"""
            <html><head><meta charset='utf-8'>
            <style>body{{font-family:Arial; font-size:11pt}} h2{{margin:0 0 12px 0}} table{{font-size:10pt}}</style>
            </head><body>
            <h2>{html_escape(subj)}</h2>
            <hr>
            <div>{body_html}</div>
            </body></html>
            """
            doc = QTextDocument()
            doc.setHtml(html)
            doc.print(printer)
            print(f"[PDF] Guardado en: {path}")
        except Exception as e:
            QMessageBox.critical(self, "PDF", f"No se pudo generar el PDF:\n{human_ex(e)}")

    def send_email(self):
        server = self.main.settings.value("accounts/smtp_server", "smtp.gmail.com", str)
        port = int(self.main.settings.value("accounts/smtp_port", 587, int))
        user = self.main.settings.value("accounts/smtp_email", "", str)
        password = self.main.settings.value("accounts/smtp_password", "", str)
        to_addr = self.mail_edit.text().strip()
        subject = self.subject_edit.text().strip()
        body_html, body_text = self._current_body_html_text()

        if not (server and port and user and password and to_addr and subject and body_text):
            QMessageBox.warning(self, "Correo", "Completa credenciales SMTP en Configuración y los campos del correo.")
            return

        try:
            from email.message import EmailMessage
            msg = EmailMessage()
            msg["From"] = user
            msg["To"] = to_addr
            msg["Subject"] = subject
            # Parte texto y parte HTML
            msg.set_content(body_text)
            msg.add_alternative(body_html, subtype="html")

            print(f"[SMTP] Conectando a {server}:{port} como {user}…")
            if port == 465:
                s = smtplib.SMTP_SSL(server, port, timeout=10)
                s.set_debuglevel(1)
            else:
                s = smtplib.SMTP(server, port, timeout=10)
                s.set_debuglevel(1)
                s.ehlo()
                if port == 587:
                    print("[SMTP] Iniciando TLS…")
                    s.starttls(); s.ehlo()

            print("[SMTP] Autenticando…")
            s.login(user, password)
            s.send_message(msg)
            s.quit()
            QMessageBox.information(self, "Correo", "Correo enviado correctamente.")
        except Exception as e:
            print("[SMTP][ERROR]", human_ex(e))
            QMessageBox.critical(self, "Correo", f"No se pudo enviar el correo:\n{human_ex(e)}")

# ------------------- Configuración -------------------

def format_bytes_gb(n: int) -> str:
    try:
        return f"{round(n/1024/1024/1024)} GB"
    except Exception:
        return "N/D"

def get_system_info() -> str:
    # CPU
    cpu = platform.processor() or os.environ.get("PROCESSOR_IDENTIFIER", "").strip() or "N/D"
    # RAM
    ram_total = "N/D"
    if psutil:
        try:
            ram_total = format_bytes_gb(psutil.virtual_memory().total)
        except Exception:
            pass
    # GPUs
    gpus = []
    try:
        # Windows: wmic
        out = subprocess.check_output(["wmic","path","win32_VideoController","get","Name,AdapterRAM"], stderr=subprocess.DEVNULL).decode(errors="ignore").splitlines()
        for ln in out[1:]:
            ln = " ".join(ln.split()).strip()
            if not ln: continue
            if "AdapterRAM" in ln: continue
            # Línea viene como: "NVIDIA GeForce RTX 3070 Laptop GPU  8589934592"
            parts = ln.rsplit(" ", 1)
            name = parts[0].strip()
            mem = ""
            if len(parts) == 2 and parts[1].isdigit():
                mem = format_bytes_gb(int(parts[1]))
            gpus.append(f"{name}{f' ({mem})' if mem else ''}")
    except Exception:
        pass
    if not gpus:
        gpus = ["No detectado"]

    # CUDA
    cuda_line = "No se detectó una GPU NVIDIA."
    try:
        out = subprocess.check_output(["nvidia-smi","--query-gpu=cuda_version","--format=csv,noheader"], stderr=subprocess.DEVNULL, timeout=2).decode().strip()
        if out:
            cuda_line = f"Versión de CUDA instalada: {out}"
    except Exception:
        # Fallback con torch
        try:
            import torch
            if torch.cuda.is_available():
                cuda_line = f"Versión de CUDA (PyTorch): {getattr(torch.version,'cuda', 'desconocida')}"
        except Exception:
            pass

    # Formato estilo tu ejemplo
    lines = []
    lines.append("[PROCESADOR]")
    lines.append(cpu)
    lines.append("")
    lines.append("[MEMORIA RAM]")
    lines.append(f"Capacidad Total: {ram_total}")
    lines.append("")
    lines.append("[TARJETA GRÁFICA (GPU)]")
    for g in gpus:
        lines.append(g)
    lines.append("")
    lines.append("[VERSIÓN DE CUDA]")
    if "NVIDIA" in " ".join(gpus):
        lines.append("Se detectó una GPU NVIDIA...")
    lines.append(cuda_line)

    return "\n".join(lines)

class SettingsDialog(QDialog):
    def __init__(self, parent_main):
        super().__init__(parent_main)
        self.main = parent_main
        self.setWindowTitle("Configuración")
        self.setMinimumSize(780, 620)
        lay = QVBoxLayout(self)

        # Topbar (solo back y título)
        top = QHBoxLayout()
        back = QPushButton("←"); back.setFixedSize(36, 28); back.clicked.connect(self.accept)
        title = QLabel("Configuración")
        f = QFont(); f.setPointSize(14); title.setFont(f)
        top.addWidget(back); top.addStretch(1); top.addWidget(title); top.addStretch(5)
        lay.addLayout(top)

        self.tabs = QTabWidget()
        lay.addWidget(self.tabs, 1)

        self._tab_general()       # incluye sistema / versión / update + rutas
        self._tab_accounts_api()  # cuentas, API y nombre/empresa

        # Lanzar validaciones automáticas
        QTimer.singleShot(400, self._auto_validate_all)

    def _tab_general(self):
        w = QWidget(); l = QVBoxLayout(w)

        # ---- Ruta TGP (carpeta)
        grp1 = QGroupBox("Ruta actual de TGP (exe/lnk)")
        g1 = QVBoxLayout(grp1)
        self.ed_prog = QLineEdit(self.main.settings.value("paths/program_path", os.getcwd(), str))
        self.ed_prog.setReadOnly(True)
        b1 = QPushButton("Elegir carpeta de destino")
        b1.clicked.connect(self._choose_prog_dir)
        g1.addWidget(self.ed_prog); g1.addWidget(b1)

        # ---- Carpeta PDF
        grp2 = QGroupBox("Ruta actual para guardar los PDF")
        g2 = QVBoxLayout(grp2)
        self.ed_pdf = QLineEdit(self.main.settings.value("paths/pdf_dir", os.getcwd(), str))
        self.ed_pdf.setReadOnly(True)
        b2 = QPushButton("Elegir carpeta de destino")
        b2.clicked.connect(self._choose_pdf_dir)
        g2.addWidget(self.ed_pdf); g2.addWidget(b2)

        # ---- Sistema / Versión
        grp_sys = QGroupBox("Características:")
        gl = QVBoxLayout(grp_sys)
        self.lbl_specs = QLabel(get_system_info())
        self.lbl_specs.setStyleSheet("font-family:Consolas, monospace; background:#f6f7fb; border:1px solid #e1e3ea; border-radius:8px; padding:8px;")
        self.lbl_specs.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        gl.addWidget(self.lbl_specs)

        # Versión + botón update grande
        hl = QVBoxLayout()
        self.lbl_version = QLabel(f"Versión: v{APP_VERSION}")
        self.btn_update = QPushButton("Comprobando actualizaciones…")
        self.btn_update.setEnabled(False)
        self.btn_update.setMinimumHeight(28)
        self.btn_update.setStyleSheet("""
            QPushButton {
                background:#e9edf7; border:1px solid #c9d3ee; color:#4f6fb5;
                font-weight:600; border-radius:6px; padding:8px;
            }
            QPushButton:disabled { color:#8ea0c5; }
        """)
        self.btn_update.clicked.connect(self._start_update)
        hl.addWidget(self.lbl_version)
        hl.addWidget(self.btn_update)

        l.addWidget(grp1); l.addWidget(grp2); l.addWidget(grp_sys); l.addLayout(hl); l.addStretch(1)
        self.tabs.addTab(w, "General")

        # Verificación de updates
        self._update_info = None
        self._upd = UpdateCheckerWorker()
        self._upd.finished.connect(self._on_update_check)
        self._upd.start()

    def _choose_prog_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Selecciona carpeta de TGP", self.ed_prog.text() or os.getcwd())
        if d:
            self.ed_prog.setText(d)
            self.main.settings.setValue("paths/program_path", d)

    def _choose_pdf_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Selecciona carpeta PDF", self.ed_pdf.text() or os.getcwd())
        if d:
            self.ed_pdf.setText(d)
            self.main.settings.setValue("paths/pdf_dir", d)

    def _tab_accounts_api(self):
        w = QWidget(); l = QVBoxLayout(w)

        # Nombre/Empresa
        grp_me = QGroupBox("Identidad (Remitente)")
        fme = QFormLayout(grp_me)
        self.ed_company = QLineEdit(self.main.settings.value("accounts/company_name", "", str))
        fme.addRow("Nombre / Empresa:", self.ed_company)

        # Google API
        grp_g = QGroupBox("Google API (Gemini)")
        fg = QFormLayout(grp_g)
        self.ed_api = QLineEdit(self.main.settings.value("accounts/google_api_key", "", str))
        self.lbl_api_state = QLabel("●"); self.lbl_api_state.setStyleSheet("color: grey; font-size:18px;")
        self.ed_api.setEchoMode(QLineEdit.EchoMode.Password)
        fg.addRow("API Key:", self.ed_api)
        fg.addRow("Estado:", self.lbl_api_state)

        # SMTP
        grp_s = QGroupBox("Correo saliente (SMTP)")
        fs = QFormLayout(grp_s)
        self.ed_smtp_server = QLineEdit(self.main.settings.value("accounts/smtp_server", "smtp.gmail.com", str))
        self.ed_smtp_port = QLineEdit(str(self.main.settings.value("accounts/smtp_port", 587, int)))
        self.ed_smtp_email = QLineEdit(self.main.settings.value("accounts/smtp_email", "", str))
        self.ed_smtp_pass = QLineEdit(self.main.settings.value("accounts/smtp_password", "", str))
        self.ed_smtp_pass.setEchoMode(QLineEdit.EchoMode.Password)
        self.lbl_smtp_state = QLabel("●"); self.lbl_smtp_state.setStyleSheet("color: grey; font-size:18px;")

        fs.addRow("Servidor:", self.ed_smtp_server)
        fs.addRow("Puerto:", self.ed_smtp_port)
        fs.addRow("Correo:", self.ed_smtp_email)
        fs.addRow("Contraseña:", self.ed_smtp_pass)
        fs.addRow("Estado:", self.lbl_smtp_state)

        l.addWidget(grp_me)
        l.addWidget(grp_g)
        l.addWidget(grp_s)
        l.addStretch(1)
        self.tabs.addTab(w, "Cuentas / API")

        # Timers automáticos
        self._t_api = QTimer(self); self._t_api.setSingleShot(True)
        self._t_api.timeout.connect(self._validate_api_auto)
        self.ed_api.textChanged.connect(lambda _=None: self._t_api.start(700))

        self._t_smtp = QTimer(self); self._t_smtp.setSingleShot(True)
        self._t_smtp.timeout.connect(self._validate_smtp_auto)
        for wdg in (self.ed_smtp_server, self.ed_smtp_port, self.ed_smtp_email, self.ed_smtp_pass):
            wdg.textChanged.connect(lambda _=None: self._t_smtp.start(800))

        # Guardar nombre/empresa al vuelo
        self.ed_company.textChanged.connect(lambda: self.main.settings.setValue("accounts/company_name", self.ed_company.text().strip()))

    # Auto-validaciones
    def _auto_validate_all(self):
        self._validate_api_auto()
        self._validate_smtp_auto()

    def _validate_api_auto(self):
        key = self.ed_api.text().strip()
        self.main.settings.setValue("accounts/google_api_key", key)
        if not key:
            print("[CFG][API] Sin clave.")
            self.lbl_api_state.setStyleSheet("color: grey; font-size:18px;")
            return
        try:
            print("[CFG][API] Validando clave…")
            genai.configure(api_key=key)
            models = list(genai.list_models())
            ok = any("generateContent" in m.supported_generation_methods and "gemini-1.5-flash" in m.name for m in models)
            print("[CFG][API] Resultado:", "OK" if ok else "FAIL")
            self.lbl_api_state.setStyleSheet("color: #2ecc71; font-size:18px;" if ok else "color: #e74c3c; font-size:18px;")
        except Exception as e:
            print("[CFG][API][ERROR]", human_ex(e))
            self.lbl_api_state.setStyleSheet("color: #e74c3c; font-size:18px;")

    def _validate_smtp_auto(self):
        server = self.ed_smtp_server.text().strip()
        try:
            port = int(self.ed_smtp_port.text().strip())
        except Exception:
            port = 0
        email = self.ed_smtp_email.text().strip()
        pwd = self.ed_smtp_pass.text()

        self.main.settings.setValue("accounts/smtp_server", server)
        self.main.settings.setValue("accounts/smtp_port", port if port else 0)
        self.main.settings.setValue("accounts/smtp_email", email)
        self.main.settings.setValue("accounts/smtp_password", pwd)

        if not (server and port and email and pwd):
            print("[CFG][SMTP] Datos incompletos.")
            self.lbl_smtp_state.setStyleSheet("color: grey; font-size:18px;")
            return
        try:
            print(f"[CFG][SMTP] Probando {server}:{port} como {email}…")
            if port == 465:
                s = smtplib.SMTP_SSL(server, port, timeout=5)
                s.set_debuglevel(1)
            else:
                s = smtplib.SMTP(server, port, timeout=5)
                s.set_debuglevel(1)
                s.ehlo()
                if port == 587:
                    print("[CFG][SMTP] STARTTLS…")
                    s.starttls(); s.ehlo()
            print("[CFG][SMTP] Login…")
            s.login(email, pwd)
            s.quit()
            print("[CFG][SMTP] OK.")
            self.lbl_smtp_state.setStyleSheet("color: #2ecc71; font-size:18px;")
        except Exception as e:
            print("[CFG][SMTP][ERROR]", human_ex(e))
            self.lbl_smtp_state.setStyleSheet("color: #e74c3c; font-size:18px;")

    # Updates
    def _on_update_check(self, info: dict):
        self._update_info = info
        if not info.get("ok"):
            self.btn_update.setText("No se pudo comprobar updates")
            self.btn_update.setEnabled(False)
            return
        tag = (info.get("tag") or "").lstrip("v")
        newer = tag and (tag != APP_VERSION)
        if newer:
            self.btn_update.setText(f"Estás en la última versión" if not newer else f"Actualizar a v{tag}")
            self.btn_update.setEnabled(True)
        else:
            self.btn_update.setText("Estás en la última versión")
            self.btn_update.setEnabled(False)

    def _start_update(self):
        info = self._update_info or {}
        assets = info.get("assets") or []
        if not assets:
            QMessageBox.warning(self, "Update", "El release no tiene assets.")
            return
        zip_asset = None
        for a in assets:
            if str(a.get("name", "")).lower().endswith(".zip"):
                zip_asset = a; break
        if not zip_asset:
            QMessageBox.warning(self, "Update", "No se encontró un .zip en el release.")
            return

        url = zip_asset["browser_download_url"]
        name = zip_asset["name"]
        tmp = resource_path("update_temp")
        os.makedirs(tmp, exist_ok=True)
        dst = os.path.join(tmp, name)
        try:
            with requests.get(url, stream=True) as r:
                r.raise_for_status()
                with open(dst, "wb") as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
            updater = resource_path("updater.py")
            if not os.path.exists(updater):
                QMessageBox.critical(self, "Update", "No se encontró updater.py en la carpeta.")
                return
            pid = os.getpid()
            new_version = info.get("tag", "").lstrip("v")
            creationflags = subprocess.DETACHED_PROCESS if sys.platform == "win32" else 0
            subprocess.Popen([sys.executable, updater, str(pid), new_version], creationflags=creationflags)
            self.accept()
            self.main.close()
        except Exception as e:
            QMessageBox.critical(self, "Update", f"Fallo al actualizar:\n{human_ex(e)}")

# ------------------- Inicio (Home) -------------------

class HomePage(QWidget):
    go_a = pyqtSignal()
    go_b = pyqtSignal()
    open_settings = pyqtSignal()
    def __init__(self):
        super().__init__()
        lay = QVBoxLayout(self)

        # Barra superior: título centrado + engranaje
        top = QHBoxLayout()
        top.addSpacing(36)  # simetría con engranaje
        title = QLabel("BALADA PACKAGING .S.L.U.")
        f = QFont(); f.setPointSize(18); f.setBold(True); title.setFont(f)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        top.addWidget(title, 1)
        gear = QPushButton("⚙️"); gear.setFixedSize(36, 28)
        gear.clicked.connect(self.open_settings.emit)
        top.addWidget(gear)
        lay.addLayout(top)

        # Dos imágenes (botones)
        imgs = QHBoxLayout()
        self.img_a = ClickLabel(); self._load_img(self.img_a, "multimedia/Opcion_A.png")
        self.img_b = ClickLabel(); self._load_img(self.img_b, "multimedia/Opcion_B.png")

        self.img_a.clicked.connect(self.go_a.emit)
        self.img_b.clicked.connect(self.go_b.emit)

        imgs.addWidget(self.img_a, 1)
        imgs.addWidget(self.img_b, 1)
        lay.addLayout(imgs, 1)

        # Etiquetas
        foot = QHBoxLayout()
        la = QLabel("Imagen A"); la.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lb = QLabel("Imagen B"); lb.setAlignment(Qt.AlignmentFlag.AlignCenter)
        foot.addWidget(la, 1); foot.addWidget(lb, 1)
        lay.addLayout(foot)

    def _load_img(self, label: QLabel, relpath: str):
        p = resource_path(relpath)
        if os.path.exists(p):
            pm = QPixmap(p)
            label.setPixmap(pm.scaled(QSize(520, 360), Qt.AspectRatioMode.KeepAspectRatio,
                                      Qt.TransformationMode.SmoothTransformation))
            label.setStyleSheet("border:1px solid #d6d6d6; border-radius:12px;")
        else:
            label.setText("Imagen no encontrada")
            label.setStyleSheet("color:#9aa0a6; border:1px dashed #d6d6d6; border-radius:12px;")
            label.setMinimumSize(520, 360)
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)

# ------------------- MainWindow ----------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings(APP_ORG, APP_NAME)

        self.setWindowTitle("BALADA PACKAGING · Presupuestos")
        self.setMinimumSize(1100, 720)

        # Stacked central
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        # Páginas
        self.page_home = HomePage()
        self.page_a = OptionAPage(self)
        self.page_b = OptionBPage(self)

        self.page_home.go_a.connect(lambda: self.stack.setCurrentWidget(self.page_a))
        self.page_home.go_b.connect(lambda: self.stack.setCurrentWidget(self.page_b))
        self.page_home.open_settings.connect(self.open_settings)

        self.stack.addWidget(self.page_home)
        self.stack.addWidget(self.page_a)
        self.stack.addWidget(self.page_b)
        self.stack.setCurrentWidget(self.page_home)

        # Sin botón “Configuración” duplicado (sólo engranaje).

        # Console update info
        self._console_update_check()

    def go_home(self):
        self.stack.setCurrentWidget(self.page_home)

    def open_settings(self):
        dlg = SettingsDialog(self)
        dlg.exec()

    def _console_update_check(self):
        def _bg():
            try:
                api_url = f"https://api.github.com/repos/{GITHUB_USER}/{GITHUB_REPO}/releases/latest"
                r = requests.get(api_url, timeout=10); r.raise_for_status()
                tag = r.json().get("tag_name", "").lstrip("v")
                print(f"[UPDATE] Local: {APP_VERSION}  —  Remoto: {tag}  —  Disponible: {bool(tag and tag != APP_VERSION)}")
            except Exception as e:
                print("[UPDATE] No se pudo comprobar:", human_ex(e))
        QTimer.singleShot(1000, _bg)

# ------------------- main() --------------------------

def main():
    icon_path = resource_path("BitStation.ico")
    app = QApplication(sys.argv)
    app.setOrganizationName(APP_ORG); app.setApplicationName(APP_NAME)
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    win = MainWindow()
    win.show()
    return app.exec()

if __name__ == "__main__":
    sys.exit(main())
