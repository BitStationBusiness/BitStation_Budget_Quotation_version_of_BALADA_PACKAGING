# BALADA PACKAGING · Presupuestos — versión con TGP integrado (auto-run)
# Python 3.10 / Windows / PyQt6
# - Cámara en Facturas (Option A) con rotación y auto-crop
# - Cámara en Presupuestos (Option B) con rotación y auto-crop
# - Tabla Markdown → HTML con estilos inline (vista previa, Gmail y PDF)
# - PDF A4 con márgenes y tipografía legible
# - Impresión como imagen (PDF→QImage 300 ppp)
# - Automatización TGP integrada (desde Option A > Procesar) **CON ventana de confirmación previa**
# - Parada de emergencia GLOBAL **solo con tecla ESC**
# ---------------------------------------------------------------------

import os
import re
import sys
import json
import time
import smtplib
import shutil
import platform
import subprocess
import tempfile
import threading
from typing import List, Dict, Any, Optional, Tuple

# Qt
from PyQt6.QtCore import (
    Qt, QSize, QTimer, QThread, pyqtSignal, QSettings, QMarginsF, QEvent
)
from PyQt6.QtGui import (
    QIcon, QPixmap, QImage, QFont, QTextDocument, QPageLayout, QPageSize, QPainter
)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QStackedWidget, QLineEdit, QTextEdit,
    QComboBox, QMessageBox, QGroupBox, QFormLayout, QTabWidget, QDialog,
    QCheckBox, QTableWidget, QTableWidgetItem, QHeaderView, QStyledItemDelegate,
    QTextBrowser
)
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog, QPrinterInfo

# IA y PDF
import fitz  # PyMuPDF
import google.generativeai as genai
from google.generativeai.types import GenerationConfig

# HTTP/updates
import requests

# Cámara
import cv2

# Sistema
try:
    import psutil
except ImportError:
    psutil = None

# Entrada global
try:
    from pynput import keyboard, mouse
except ImportError:
    keyboard = None
    mouse = None

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

# ---- HTML helpers ----
def html_escape(s: str) -> str:
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )

def markdown_table_to_html(text: str) -> Optional[str]:
    """Convierte todo el cuerpo en HTML (párrafos + tablas Markdown)."""
    lines = text.splitlines()
    i = 0  # [FIX] corregido (antes había un '>' accidental que causaba SyntaxError)
    out_parts: List[str] = []
    found_table = False

    def emit_paragraphs(chunk: List[str]):
        for ln in chunk:
            if not ln.strip():
                out_parts.append("<p style='margin:0 0 12px'>&nbsp;</p>")
            else:
                out_parts.append(f"<p style='margin:0 0 10px'>{html_escape(ln)}</p>")

    while i < len(lines):
        if lines[i].strip().startswith("|"):
            block = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                block.append(lines[i].strip())
                i += 1

            cleaned = []
            for raw in block:
                parts = [c.strip() for c in raw.split("|")]
                if parts and parts[0] == "":
                    parts = parts[1:]
                if parts and parts[-1] == "":
                    parts = parts[:-1]
                if parts and set("".join(parts)) <= set("-: "):
                    continue
                if parts:
                    cleaned.append(parts)

            if cleaned:
                found_table = True
                thead_cells = "".join(
                    f"<th style='background:#f6f8fa;border:1px solid #dfe3e8;"
                    f"padding:8px 10px;text-align:left;font-weight:600'>{html_escape(c)}</th>"
                    for c in cleaned[0]
                )
                body_rows = []
                for row in cleaned[1:]:
                    tds = "".join(
                        f"<td style='border:1px solid #e6e9ef;padding:8px 10px;vertical-align:top'>{html_escape(c)}</td>"
                        for c in row
                    )
                    body_rows.append(f"<tr style='page-break-inside:avoid'>{tds}</tr>")

                table_html = (
                    "<table style='border-collapse:collapse;width:100%;table-layout:fixed;margin:8px 0 12px;"
                    "font-family:Arial;font-size:12pt;color:#222'>"
                    f"<thead><tr>{thead_cells}</tr></thead>"
                    f"<tbody>{''.join(body_rows)}</tbody></table>"
                )
                out_parts.append(table_html)
            else:
                emit_paragraphs(block)
        else:
            non = []
            while i < len(lines) and not lines[i].strip().startswith("|"):
                non.append(lines[i])
                i += 1
            emit_paragraphs(non)

    html = "".join(out_parts).strip()
    return html if found_table else None

# ------------------- PDF → QImages (raster a 300 ppp) -------------------

def render_pdf_to_qimages(pdf_path: str, dpi: int = 300) -> List[QImage]:
    doc = fitz.open(pdf_path)
    images: List[QImage] = []
    try:
        for page in doc:
            pix = page.get_pixmap(dpi=dpi, alpha=False)
            img = QImage(
                pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888
            )
            images.append(img.copy())
    finally:
        doc.close()
    return images

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

    def _placeholder(self):
        self.setPixmap(QPixmap())
        self.setText("Arrastra una imagen/PDF aquí")

    def reset_placeholder(self):
        self._pix = None
        self.bytes = None
        self.mime = None
        self.loaded_path = None
        self._placeholder()

    def show_qimage(self, qimg: QImage):
        self._pix = QPixmap.fromImage(qimg)
        self.setText("")
        self._rescale()

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

class EnterKeyDelegate(QStyledItemDelegate):
    """Intercepta Enter/Return en editores de celdas y delega en la tabla."""
    def createEditor(self, parent, option, index):
        editor = super().createEditor(parent, option, index)
        try:
            editor.installEventFilter(self)
        except Exception:
            pass
        return editor

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.KeyPress:
            try:
                key = event.key()
            except Exception:
                key = None
            if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
                table = self.parent()
                if hasattr(table, "handle_enter_action"):
                    table.handle_enter_action()
                    return True
        return super().eventFilter(obj, event)

class EditableTable(QTableWidget):
    """Enter:
       - Si la fila actual está vacía ⇒ se elimina.
       - Si tiene contenido ⇒ inserta una fila debajo y entra en edición.
       Al cambiar de fila: si la anterior queda vacía ⇒ se elimina."""
    def __init__(self, rows: int, cols: int, parent=None):
        super().__init__(rows, cols, parent)
        self._cols = cols

    def row_all_blank(self, r: int) -> bool:
        if r < 0 or r >= self.rowCount():
            return False
        for c in range(self._cols):
            it = self.item(r, c)
            if it and it.text().strip():
                return False
        return True

    def ensure_items(self, r: int):
        for c in range(self._cols):
            if not self.item(r, c):
                self.setItem(r, c, QTableWidgetItem(""))

    def handle_enter_action(self):
        r = self.currentRow()
        if r < 0:
            insert_at = self.rowCount()
            self.insertRow(insert_at)
            self.ensure_items(insert_at)
            self.setCurrentCell(insert_at, 0)
            self.editItem(self.item(insert_at, 0))
            return

        if self.row_all_blank(r):
            self.removeRow(r)
            if self.rowCount() > 0:
                self.setCurrentCell(min(r, self.rowCount() - 1), 0)
            return

        insert_at = r + 1
        self.insertRow(insert_at)
        self.ensure_items(insert_at)
        self.setCurrentCell(insert_at, 0)
        self.editItem(self.item(insert_at, 0))

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.handle_enter_action()
            return
        super().keyPressEvent(e)

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
        return {}

# ------------------- AUTOMATIZACIÓN TGP -------------------

class TgpAutomationWorker(QThread):
    """Abre/cierra TGP, login 'visa', Clicks 2.1 → **espera 5s** → (VENTANA de confirmación) →
    Clicks 2.2 (Ctrl+B x veces) → traspaso de tabla → pregunta impresión.
    Parada de emergencia con ESC en cualquier momento.
    """

    # Señales hacia la UI
    log = pyqtSignal(str)
    need_user_pause = pyqtSignal(str)    # diálogo de pausa ("Aceptar / Cancelar")
    ask_print = pyqtSignal()             # pedir decisión de impresión
    show_info = pyqtSignal(str)          # info final
    finished_ok = pyqtSignal()
    aborted = pyqtSignal(str)

    def __init__(self, tgp_path: str, rows_to_transfer: List[Dict[str, str]]):
        super().__init__()
        self.tgp_path = tgp_path
        self.rows = rows_to_transfer or []

        # Flags de control
        self._pause_flag = threading.Event()
        self._resume_flag = threading.Event()
        self._cancel_flag = threading.Event()

        # Decisión de impresión
        self._print_event = threading.Event()
        self._print_choice: Optional[bool] = None

        # Supresión de pausa mientras inyectamos input
        self._suppress_until = 0.0

    # ---------- Control desde el listener global ----------
    def _now(self) -> float:
        import time as _t
        return _t.time()

    def _suppress_for(self, seconds: float):
        self._suppress_until = max(self._suppress_until, self._now() + max(0.0, seconds))

    def is_inputting(self) -> bool:
        return self._now() < self._suppress_until

    # [NEW] Cancelación inmediata (emergencia ESC)
    def cancel_now(self):  # llamado desde listener
        self._cancel_flag.set()
        self._resume_flag.set()  # por si está en pausa/espera
        self.log.emit("[EMERGENCIA] Cancelación inmediata solicitada (ESC).")

    def request_global_pause(self):
        """(Compatibilidad; no usamos una pausa manual distinta a ESC)"""
        self._pause_flag.set()
        self._resume_flag.clear()
        self.need_user_pause.emit("user-request")

    def continue_after_pause(self):
        self._pause_flag.clear()
        self._resume_flag.set()

    def cancel_after_pause(self):
        self._cancel_flag.set()
        self._resume_flag.set()

    # Decisión de impresión desde la UI
    def set_print_decision(self, do_print: bool):
        self._print_choice = bool(do_print)
        self._print_event.set()

    # ---------- Helpers de input ----------
    def _check_cancel(self):  # [NEW]
        if self._cancel_flag.is_set():
            raise RuntimeError("Cancelado por el usuario.")

    def _sleep(self, sec: float):
        import time as _t
        # micro-sleeps para responder rápido a cancelaciones
        end = _t.time() + max(0.0, sec)
        while _t.time() < end:
            self._check_cancel()
            _t.sleep(0.02)

    def _ensure_ctrls(self):
        if keyboard is None or mouse is None:
            raise RuntimeError("Falta 'pynput'. Instálalo para usar la automatización de TGP.")
        self._kbd = keyboard.Controller()
        self._mouse = mouse.Controller()

    def _type_text(self, text: str, key_delay: float = 0.01):
        """Escribe texto con supresión temporal de la pausa global."""
        self._suppress_for(max(0.15, len(text) * (key_delay + 0.002)))
        for ch in text:
            self._check_cancel()  # [NEW]
            self._kbd.type(ch)
            self._sleep(key_delay)

    def _press_key(self, k, times: int = 1, delay: float = 0.05):
        self._suppress_for(max(0.15, times * (delay + 0.02)))
        for _ in range(times):
            self._check_cancel()  # [NEW]
            self._kbd.press(k)
            self._kbd.release(k)
            self._sleep(delay)

    def _hotkey_ctrl_b(self):
        self._check_cancel()  # [NEW]
        self._suppress_for(0.20)
        self._kbd.press(keyboard.Key.ctrl)
        self._kbd.press('b')
        self._kbd.release('b')
        self._kbd.release(keyboard.Key.ctrl)
        self._sleep(0.06)

    def _click_xy(self, x: int, y: int, btn=None):
        self._check_cancel()  # [NEW]
        btn = btn or mouse.Button.left
        self._suppress_for(0.20)
        self._mouse.position = (int(x), int(y))
        self._sleep(0.05)
        self._mouse.click(btn, 1)
        self._sleep(0.15)

    def _wait_or_abort(self, sec: float):
        """Espera pasiva, pero aborta si el usuario canceló."""
        step = 0.1
        t = 0.0
        while t < sec:
            self._check_cancel()
            self._sleep(step)
            t += step

    # ---------- Gestión de proceso TGP ----------
    def _kill_existing_tgp(self):
        if not psutil:
            return
        targets = ("tgp", "tgpro", "tgprofesional", "tg profesional")
        for p in psutil.process_iter(attrs=["pid", "name", "exe"]):
            try:
                name = (p.info.get("name") or "").lower()
                exe = (p.info.get("exe") or "").lower()
                if any(t in name for t in targets) or any(t in exe for t in targets):
                    self.log.emit(f"[TGP] Cerrando proceso previo: PID={p.pid}, name={p.info.get('name')}")
                    p.terminate()
            except Exception:
                continue
        self._wait_or_abort(1.5)

    def _launch_tgp(self):
        self.log.emit("[TGP] Abriendo aplicación…")
        try:
            if self.tgp_path.lower().endswith(".lnk") and sys.platform.startswith("win"):
                os.startfile(self.tgp_path)  # type: ignore
            else:
                creationflags = subprocess.DETACHED_PROCESS if sys.platform == "win32" else 0
                subprocess.Popen([self.tgp_path], creationflags=creationflags)
        except Exception as e:
            raise RuntimeError(f"No se pudo iniciar TGP: {human_ex(e)}")
        self._wait_or_abort(3.5)

    def _login_with_visa(self):
        # [FIX] respeta tu timing para evitar que los clicks siguientes se encimen
        self.log.emit("[TGP] Login: escribiendo 'visa' y ENTER (x2)…")
        self._type_text("visa", key_delay=0.02)
        self._wait_or_abort(0.35)
        self._press_key(keyboard.Key.enter, times=2, delay=0.12)
        self._wait_or_abort(3.0)

    # ---------- Secuencias Clicks ----------
    def _do_clicks_2_1(self):
        """Clicks 2.1 con 500ms entre clics."""
        self.log.emit("[TGP] Clicks 2.1…")
        self._click_xy(427, 102); self._wait_or_abort(0.5)
        self._click_xy(86, 161);  self._wait_or_abort(0.5)
        self._click_xy(98, 174);  self._wait_or_abort(0.5)

    def _do_clicks_2_2(self):
        self.log.emit("[TGP] Clicks 2.2 (Ctrl+B + Enter x25)…")
        self._click_xy(23, 342)  # foco en la grilla
        self._wait_or_abort(0.5)
        for _ in range(25):
            self._check_cancel()  # [NEW]
            self._hotkey_ctrl_b()
            self._press_key(keyboard.Key.enter, times=1, delay=0.06)

    # ---------- Pausa guiada ----------
    def _pause_for_user_to_pick_client(self):
        self.log.emit("[TGP] Pausa: selecciona cliente/pedido y pulsa Aceptar.")
        self._pause_flag.set()
        self._resume_flag.clear()
        self.need_user_pause.emit("select-client")
        while not self._resume_flag.is_set():
            self._check_cancel()  # [NEW]
            self._sleep(0.1)
        self._pause_flag.clear()
        self.log.emit("[TGP] Reanudando…")
        self._wait_or_abort(0.3)

    # ---------- Traspaso tabla a TGP ----------
    def _focus_description_field(self):
        self.log.emit("[TGP] Foco en 'Descripción' (221,343)…")
        self._click_xy(221, 343)
        self._wait_or_abort(0.25)

    def _transfer_all_rows(self):
        if not self.rows:
            self.log.emit("[TGP] No hay filas para transferir.")
            return
        self._focus_description_field()

        for idx, r in enumerate(self.rows, start=1):
            self._check_cancel()  # [NEW]
            desc = (r.get("descripcion") or "").strip()
            qty  = (r.get("cantidad") or "").strip()
            prc  = (r.get("precio") or "").strip()
            self.log.emit(f"[TGP] Fila {idx}: DESC='{desc}' | CANT='{qty}' | PRECIO='{prc}'")

            # Descripción
            if desc:
                self._type_text(desc, key_delay=0.01)
            self._press_key(keyboard.Key.enter, times=1, delay=0.08)

            # → Cantidad
            self._press_key(keyboard.Key.right, times=1, delay=0.06)
            if qty:
                self._type_text(qty, key_delay=0.01)
                self._press_key(keyboard.Key.enter, times=1, delay=0.08)  # solo si hay cantidad

            # →→ Precio
            self._press_key(keyboard.Key.right, times=2, delay=0.06)
            if prc:
                self._type_text(prc, key_delay=0.01)
                self._press_key(keyboard.Key.enter, times=1, delay=0.08)  # solo si hay precio

            # ↓ nueva fila y volver a Descripción con ← ← ←
            self._press_key(keyboard.Key.down, times=1, delay=0.08)
            self._press_key(keyboard.Key.left, times=3, delay=0.06)

        self.log.emit("[TGP] Transferencia completada.")

    # ---------- Decisión y secuencia de impresión ----------
    def _ask_print_decision(self) -> bool:
        self._print_event.clear()
        self._print_choice = None
        self.ask_print.emit()  # la UI mostrará el diálogo y llamará a set_print_decision()
        while not self._print_event.is_set():
            self._check_cancel()  # [NEW]
            self._sleep(0.1)
        return bool(self._print_choice)

        # Dentro de la clase TgpAutomationWorker
    def _do_print_sequence(self):
        """
        Secuencia de impresión con CUATRO clics (el 1.º es el nuevo):
        (518,298) -> (505,57) -> (735,384) -> (758,685)
        500 ms entre clics para evitar solapamientos.
        """
        if mouse is None:
            raise RuntimeError("Falta 'pynput' (mouse) para usar la impresión automatizada.")

        # 1) Click previo solicitado
        self._click_xy(26, 341, btn=mouse.Button.left)
        self._wait_or_abort(0.5)
        
        self._click_xy(26, 341, btn=mouse.Button.left)
        self._wait_or_abort(0.5)

        # 2) Click 1 original
        self._click_xy(505, 57)
        self._wait_or_abort(0.5)

        # 3) Click 2 original
        self._click_xy(735, 384)
        self._wait_or_abort(0.5)

        # 4) Click 3 original
        self._click_xy(758, 685)
        self._wait_or_abort(0.5)


    # ---------- Orquestación principal ----------
    def run(self):
        try:
            self._ensure_ctrls()

            # 1) Cerrar TGP si estaba abierto
            self._kill_existing_tgp(); self._check_cancel()

            # 2) Abrir TGP
            self._launch_tgp(); self._check_cancel()

            # 3) Login con 'visa'
            self._login_with_visa(); self._check_cancel()

            # 4) Clicks 2.1
            self._do_clicks_2_1(); self._check_cancel()

            # 6) ⚠️ Ventana de confirmación ANTES del bucle Ctrl+B
            self._pause_for_user_to_pick_client(); self._check_cancel()

            # 7) Clicks 2.2 (post-confirmación)
            self._do_clicks_2_2(); self._check_cancel()

            # 8) Traspaso de la tabla
            self._transfer_all_rows(); self._check_cancel()

            # 9) Preguntar si desea imprimir y ejecutar la secuencia de impresión
            if self._ask_print_decision():
                self._do_print_sequence()
                self.show_info.emit("Se completó con éxito.")

            # 10) Fin
            self.finished_ok.emit()

        except Exception as e:
            self.aborted.emit(human_ex(e))

# ------------------- Opción A (Extractor + TGP) ----------

class OptionAPage(QWidget):
    """Extractor de datos de facturas + Automatización TGP (auto-run)."""
    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        self.api_key: Optional[str] = None

        self.bytes: Optional[bytes] = None
        self.mime: Optional[str] = None
        self.src_path: Optional[str] = None
        self.pdf_rows: Optional[List[Dict[str, str]]] = None
        self._populating = False

        # TGP runtime
        self._tgp_worker: Optional[TgpAutomationWorker] = None
        self._tgp_listener = None  # keyboard.Listener

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

    # ----------- utilidades tabla ----------
    def _collect_rows_from_table(self) -> List[Dict[str, str]]:
        rows: List[Dict[str, str]] = []
        for r in range(self.table.rowCount()):
            d = (self.table.item(r, 0).text() if self.table.item(r, 0) else "").strip()
            c = (self.table.item(r, 1).text() if self.table.item(r, 1) else "").strip()
            p = (self.table.item(r, 2).text() if self.table.item(r, 2) else "").strip()
            if d or c or p:
                rows.append({"descripcion": d, "cantidad": c, "precio": p})
        return rows

    # ----------- TGP helpers (auto-run) -----------
    def _start_global_listener(self):
        if keyboard is None:
            print("[TGP] Advertencia: no se pudo iniciar listener global (instala 'pynput').")
            return

        # [FIX] ESC ahora cancela inmediatamente (no pausa)
        def on_press(k):
            if self._tgp_worker:
                try:
                    if k == keyboard.Key.esc:
                        print("[EMERGENCIA] ESC detectado → cancel_now()")
                        self._tgp_worker.cancel_now()
                except Exception:
                    pass

        self._tgp_listener = keyboard.Listener(on_press=on_press)
        self._tgp_listener.start()
        print("[EMERGENCIA] Listener GLOBAL activo. Pulsa ESC para CANCELAR de inmediato.")

    def _stop_global_listener(self):
        if self._tgp_listener:
            try:
                self._tgp_listener.stop()
            except Exception:
                pass
            self._tgp_listener = None

    def _on_tgp_pause_dialog(self, _reason: str):
        # Ventana de confirmación ANTES del bucle Ctrl+B
        m = QMessageBox(self)
        m.setIcon(QMessageBox.Icon.Question)
        m.setWindowTitle("Confirmación — Selecciona cliente")
        m.setText(
            "Selecciona el **cliente/pedido** en TGP.\n\n"
            "Cuando estés listo, pulsa **Aceptar** para continuar.\n"
            "O pulsa **Cancelar** para abortar la secuencia."
        )
        aceptar = m.addButton("Aceptar", QMessageBox.ButtonRole.AcceptRole)
        cancelar = m.addButton("Cancelar", QMessageBox.ButtonRole.RejectRole)
        m.exec()
        if self._tgp_worker:
            if m.clickedButton() == aceptar:
                self._tgp_worker.continue_after_pause()
            else:
                self._tgp_worker.cancel_after_pause()

    def _on_tgp_ask_print(self):
        m = QMessageBox(self)
        m.setIcon(QMessageBox.Icon.Question)
        m.setWindowTitle("¿Imprimir?")
        m.setText("La tabla se traspasó correctamente.\n\n¿Quieres imprimir ahora?")
        b_print = m.addButton("Imprimir", QMessageBox.ButtonRole.AcceptRole)
        b_cancel = m.addButton("Cancelar", QMessageBox.ButtonRole.RejectRole)
        m.exec()
        if self._tgp_worker:
            self._tgp_worker.set_print_decision(m.clickedButton() == b_print)

    # ----------- UI -----------

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

        self.table = EditableTable(0, 3)
        self.table.setHorizontalHeaderLabels(["Descripción", "Cantidad", "Precio"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(True)
        self.table.setItemDelegate(EnterKeyDelegate(self.table))
        self.table.currentCellChanged.connect(self._on_current_cell_changed)

        content.addWidget(self.drop, 1)
        content.addWidget(self.table, 1)
        root.addLayout(content, 1)

        # Cámara / barra inferior
        self._cap = None
        self._cam_timer = QTimer(self); self._cam_timer.setInterval(33)
        self._cam_timer.timeout.connect(self._update_preview_A)
        self._last_frame_a = None
        self._rot_a = 0

        footer = QHBoxLayout()
        self.cam_combo_a = QComboBox()
        self.cam_combo_a.addItem("Ninguna (arrastrar imagen)", -1)
        self.btn_rot_l_a = QPushButton("↺"); self.btn_rot_l_a.setFixedWidth(30)
        self.btn_rot_r_a = QPushButton("↻"); self.btn_rot_r_a.setFixedWidth(30)
        self.btn_rot_l_a.clicked.connect(lambda: self._set_rot_a(-90))
        self.btn_rot_r_a.clicked.connect(lambda: self._set_rot_a(+90))
        self.chk_autocrop_a = QCheckBox("Auto-crop")
        self.chk_autocrop_a.setChecked(True)
        self.chk_autocrop_a.stateChanged.connect(lambda *_: self._render_preview_a())
        self.btn_capture_a = QPushButton("Capturar imagen")
        self.btn_capture_a.clicked.connect(self._capture_frame_A)
        self._populate_cameras_A()
        self.cam_combo_a.currentIndexChanged.connect(self._on_cam_changed_A)

        self.btn_open = QPushButton("Abrir archivo…")
        self.btn_open.clicked.connect(self._open_file)

        # Procesar = AUTOMATIZACIÓN TGP
        self.btn_proc = QPushButton("Procesar")
        self.btn_proc.clicked.connect(self._run_tgp_automation)

        footer.addWidget(self.cam_combo_a, 3)
        footer.addSpacing(6)
        footer.addWidget(self.btn_rot_l_a); footer.addWidget(self.btn_rot_r_a)
        footer.addSpacing(8)
        footer.addWidget(self.chk_autocrop_a)
        footer.addSpacing(8)
        footer.addWidget(self.btn_capture_a)
        footer.addStretch(1)
        footer.addWidget(self.btn_open)
        footer.addWidget(self.btn_proc)
        root.addLayout(footer)

        # Timer para tomar API de settings
        self._api_timer = QTimer(self); self._api_timer.setInterval(800); self._api_timer.setSingleShot(True)
        self._api_timer.timeout.connect(self._refresh_api_key)
        self._api_timer.start()

    # <<< Método propio (no anidado) >>>
    def _run_tgp_automation(self):
        # Ruta desde Configuración
        tgp_path = self.main.settings.value("paths/tgp_path", "", str).strip()
        if not tgp_path or not os.path.exists(tgp_path):
            QMessageBox.warning(self, "TGP", "Configura la ruta del .exe/.lnk en Configuración → General.")
            return
        if self._tgp_worker and self._tgp_worker.isRunning():
            QMessageBox.information(self, "TGP", "Ya hay una secuencia en ejecución.")
            return

        # Recolectar filas de la tabla (para el traspaso)
        rows = self._collect_rows_from_table()
        if not rows:
            ok = QMessageBox.question(
                self, "TGP — Sin filas",
                "La tabla está vacía. ¿Deseas continuar de todos modos?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if ok != QMessageBox.StandardButton.Yes:
                return

        # Listener global (ESC = cancelación inmediata)
        self._start_global_listener()

        # Worker
        self._tgp_worker = TgpAutomationWorker(tgp_path, rows)
        self._tgp_worker.log.connect(lambda s: print(s.strip()))
        self._tgp_worker.need_user_pause.connect(self._on_tgp_pause_dialog)
        self._tgp_worker.ask_print.connect(self._on_tgp_ask_print)
        self._tgp_worker.show_info.connect(lambda msg: QMessageBox.information(self, "TGP", msg))
        self._tgp_worker.finished_ok.connect(lambda: QMessageBox.information(self, "TGP", "Secuencia finalizada."))
        # [FIX] Mensaje emergente al cancelar
        self._tgp_worker.aborted.connect(lambda msg: QMessageBox.critical(self, "TGP", f"Secuencia cancelada: {msg}"))
        self._tgp_worker.finished.connect(self._stop_global_listener)
        self._tgp_worker.start()

    # ---------- IA (mantiene pipeline) ----------
    def _refresh_api_key(self):
        key = self.main.settings.value("accounts/google_api_key", "", str).strip()
        self.api_key = key or None
        print(f"[TGP] API Key cargada: {'sí' if self.api_key else 'no'}")

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
        self._populating = True
        try:
            self.table.setRowCount(0)
            for obj in rows:
                r = self.table.rowCount()
                self.table.insertRow(r)
                self.table.setItem(r, 0, QTableWidgetItem(str(obj.get("descripcion", ""))))
                self.table.setItem(r, 1, QTableWidgetItem(str(obj.get("cantidad", ""))))
                self.table.setItem(r, 2, QTableWidgetItem(str(obj.get("precio", ""))))
        finally:
            self._populating = False

    def _on_current_cell_changed(self, r: int, c: int, prev_r: int, prev_c: int):
        if self._populating:
            return
        if prev_r is not None and prev_r >= 0 and prev_r < self.table.rowCount():
            if self.table.row_all_blank(prev_r):
                try:
                    self.table.currentCellChanged.disconnect(self._on_current_cell_changed)
                    self.table.removeRow(prev_r)
                finally:
                    self.table.currentCellChanged.connect(self._on_current_cell_changed)

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

        schema_a = {"type": "OBJECT","properties": {"row_count": {"type": "INTEGER"}},"required": ["row_count"]}
        cfg_a = GenerationConfig(response_mime_type="application/json", response_schema=schema_a, temperature=0.0)
        resp_a = model.generate_content([self.prompt_count_a, part], generation_config=cfg_a)
        raw_a = getattr(resp_a, "text", "") or "{}"
        try:
            n_a = int(safe_json_loads(raw_a).get("row_count", 0))
        except Exception:
            n_a = 0

        schema_b = {
            "type": "OBJECT",
            "properties": {
                "lines": {
                    "type": "ARRAY",
                    "items": {
                        "type": "OBJECT",
                        "properties": {"y": {"type": "NUMBER"}, "text": {"type": "STRING"}},
                        "required": ["y", "text"]
                    }
                }
            },
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
                "rows": {
                    "type": "ARRAY",
                    "items": {
                        "type": "OBJECT",
                        "properties": {
                            "descripcion": {"type": "STRING"},
                            "cantidad": {"type": "STRING"},
                            "precio": {"type": "STRING"}
                        },
                        "required": ["descripcion", "cantidad", "precio"]
                    }
                }
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
        clean = [{
            "descripcion": str(r.get("descripcion", "")),
            "cantidad": first_number_or_blank(r.get("cantidad", "")),
            "precio": first_number_or_blank(r.get("precio", "")),
        } for r in rows]
        return clean

    def _repair_to_n(self, rows: List[Dict[str, str]], n: int) -> List[Dict[str, str]]:
        genai.configure(api_key=self.api_key)
        model = genai.GenerativeModel("gemini-1.5-flash-latest")
        schema = {
            "type": "OBJECT",
            "properties": {
                "rows": {
                    "type": "ARRAY",
                    "items": {
                        "type": "OBJECT",
                        "properties": {
                            "descripcion": {"type": "STRING"},
                            "cantidad": {"type": "STRING"},
                            "precio": {"type": "STRING"}
                        },
                        "required": ["descripcion", "cantidad", "precio"]
                    }
                }
            },
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
        fixed = [{
            "descripcion": str(r.get("descripcion", "")),
            "cantidad": first_number_or_blank(r.get("cantidad", "")),
            "precio": first_number_or_blank(r.get("precio", "")),
        } for r in data.get("rows", [])]
        if len(fixed) < n:
            fixed += [{"descripcion": "", "cantidad": "", "precio": ""} for _ in range(n - len(fixed))]
        if len(fixed) > n:
            fixed = fixed[:n]
        return fixed

    # ---------- Cámara: helpers Option A ----------
    def _populate_cameras_A(self):
        while self.cam_combo_a.count() > 1:
            self.cam_combo_a.removeItem(1)

        found = 0
        backends = [cv2.CAP_DSHOW, cv2.CAP_MSMF]
        for idx in range(0, 6):
            for be in backends:
                cap = cv2.VideoCapture(idx, be)
                if cap.isOpened():
                    ok, _ = cap.read()
                    cap.release()
                    if ok:
                        self.cam_combo_a.addItem(f"Cámara {idx}", idx)
                        print(f"[Cam-A] Detectada cámara #{idx}")
                        found += 1
                        break
                else:
                    cap.release()
        if hasattr(self, "btn_capture_a"):
            self.btn_capture_a.setEnabled(found > 0)
        print(f"[Cam-A] Total cámaras: {found}")

    def _on_cam_changed_A(self, _i: int):
        cam_index = self.cam_combo_a.currentData()
        self._stop_preview_A()
        if cam_index is None or int(cam_index) == -1:
            self.drop.reset_placeholder()
        else:
            self._start_preview_A(int(cam_index))

    def _start_preview_A(self, index: int):
        self._cap = cv2.VideoCapture(index, cv2.CAP_DSHOW)
        if not self._cap.isOpened():
            QMessageBox.critical(self, "Cámara", "No se pudo abrir la cámara seleccionada.")
            self._cap = None
            return
        self._cam_timer.start()

    def _stop_preview_A(self):
        self._cam_timer.stop()
        if self._cap:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None

    def _set_rot_a(self, delta: int):
        self._rot_a = (self._rot_a + delta) % 360
        self._render_preview_a()

    def _render_preview_a(self):
        if self._last_frame_a is None:
            return
        frame = self._last_frame_a.copy()
        if self._rot_a == 90:
            frame = cv2.rotate(frame, cv2.ROTATE_90_CLOCKWISE)
        elif self._rot_a == 180:
            frame = cv2.rotate(frame, cv2.ROTATE_180)
        elif self._rot_a == 270:
            frame = cv2.rotate(frame, cv2.ROTATE_90_COUNTERCLOCKWISE)

        if self.chk_autocrop_a.isChecked() and self.drop.width() > 0 and self.drop.height() > 0:
            h, w = frame.shape[:2]
            target_aspect = self.drop.width() / max(1, self.drop.height())
            frame_aspect = w / max(1, h)
            if abs(frame_aspect - target_aspect) > 1e-3:
                if frame_aspect > target_aspect:
                    new_w = int(h * target_aspect)
                    x0 = max(0, (w - new_w) // 2)
                    frame = frame[:, x0:x0 + new_w]
                else:
                    new_h = int(w / target_aspect)
                    y0 = max(0, (h - new_h) // 2)
                    frame = frame[y0:y0 + new_h, :]

        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb.shape
        qimg = QImage(rgb.data, w, h, ch * w, QImage.Format.Format_RGB888)
        self.drop.show_qimage(qimg)

    def _update_preview_A(self):
        if not self._cap:
            return
        ok, frame = self._cap.read()
        if not ok:
            return
        self._last_frame_a = frame
        self._render_preview_a()

    def _capture_frame_A(self):
        cam_index = self.cam_combo_a.currentData()
        if cam_index is None or cam_index == -1:
            QMessageBox.warning(self, "Cámara", "Selecciona una cámara distinta de 'Ninguna'.")
            return
        if self._last_frame_a is None:
            QMessageBox.warning(self, "Cámara", "Aún no hay imagen de la cámara.")
            return
        frame = self._last_frame_a.copy()
        if self._rot_a == 90:
            frame = cv2.rotate(frame, cv2.ROTATE_90_CLOCKWISE)
        elif self._rot_a == 180:
            frame = cv2.rotate(frame, cv2.ROTATE_180)
        elif self._rot_a == 270:
            frame = cv2.rotate(frame, cv2.ROTATE_90_COUNTERCLOCKWISE)

        tmp = resource_path("_captura_factura_tmp.png")
        cv2.imwrite(tmp, frame)
        self.drop.load_path(tmp)
        try:
            os.remove(tmp)
        except Exception:
            pass

    def stop_camera(self):
        self._stop_preview_A()

# ------------------- IA worker (Opción B) ------------

class EmailGenWorker(QThread):
    finished = pyqtSignal(str, str)
    failed = pyqtSignal(str)
    def __init__(self, api_key: str, recipient_name: str, img_bytes: bytes, mime: str):
        super().__init__()
        self.api_key = api_key
        self.recipient_name = recipient_name.strip()
        self.img_bytes = img_bytes
        self.mime = mime or "image/png"
    def run(self):
        try:
            genai.configure(api_key=self.api_key)
            model = genai.GenerativeModel("gemini-1.5-flash-latest")
            schema = {
                "type": "OBJECT",
                "properties": {"subject": {"type": "STRING"}, "body": {"type": "STRING"}},
                "required": ["subject", "body"]
            }
            cfg = GenerationConfig(response_mime_type="application/json", response_schema=schema, temperature=0.2)
            prompt = (
                "Lee el texto manuscrito/imprimido de la imagen y genera un correo de solicitud de cotización.\n"
                f"- Destinatario: {self.recipient_name or '[Nombre del destinatario]'}\n"
                "- Usa una tabla estilo Markdown con tuberías para listar ítems (| Ítem | Cantidad | Detalle |).\n"
                "- Devuelve un asunto breve y un cuerpo formal, en español.\n"
                "- No inventes datos; si faltan, deja texto genérico.\n"
                "Devuelve SOLO JSON con 'subject' y 'body'."
            )
            part = {"mime_type": self.mime, "data": self.img_bytes}
            resp = model.generate_content([prompt, part], generation_config=cfg)
            raw = getattr(resp, "text", "") or "{}"
            data = safe_json_loads(raw)
            subject = (data.get("subject") or "").strip() or "Solicitud de cotización"
            body = (data.get("body") or "").strip()
            if self.recipient_name:
                body = body.replace("[Nombre del destinatario]", self.recipient_name).replace("[Destinatario]", self.recipient_name)
            self.finished.emit(subject, body)
        except Exception as e:
            self.failed.emit(human_ex(e))

# ------------------- Opción B (Presupuestos) ---------

class OptionBPage(QWidget):
    def __init__(self, parent_main):
        super().__init__()
        self.main = parent_main
        self._ai_worker: Optional[EmailGenWorker] = None

        # cámara (live)
        self._cap: Optional[cv2.VideoCapture] = None
        self._cam_timer = QTimer(self)
        self._cam_timer.setInterval(33)  # ~30 FPS
        self._cam_timer.timeout.connect(self._update_preview)
        self._last_frame = None
        self._rot = 0

        # buffer del cuerpo (texto plano/markdown)
        self._body_raw: str = ""

        self._build_ui()

    def _compose_html_document(self, subj: str, body_text: str) -> str:
        body_html = self._body_text_to_html(body_text)
        return f"""
        <html><head><meta charset='utf-8'>
        <style>
            body{{font-family:Arial,Helvetica,sans-serif; font-size:12pt; color:#222; line-height:1.6;}}
            h2{{margin:0 0 12px; font-size:16pt;}}
            hr{{border:none; border-top:1px solid #ddd; margin:10px 0 14px;}}
            table{{border-collapse:collapse; width:100%; table-layout:fixed;}}
            thead th{{background:#f6f8fa;border:1px solid #dfe3e8;padding:6px 8px;text-align:left;font-weight:600}}
            tbody td{{border:1px solid #e6e9ef;padding:6px 8px;vertical-align:top}}
            tr, td, th{{page-break-inside:avoid;}}
            p{{margin:0 0 10px}}
        </style>
        </head><body>
        <h2>{html_escape(subj)}</h2>
        <hr/>
        {body_html}
        </body></html>
        """.strip()

    def print_preview(self):
        """Genera PDF temporal → rasteriza a 300 ppp → imprime como imagen."""
        try:
            subj = (self.subject_edit.text().strip() or "Solicitud de cotización")
            html = self._compose_html_document(subj, self._body_raw or self.body_edit.toPlainText())

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpf:
                tmp_pdf = tmpf.name
            try:
                pdf_printer = QPrinter(QPrinter.PrinterMode.HighResolution)
                pdf_printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
                pdf_printer.setOutputFileName(tmp_pdf)
                pdf_printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
                pdf_printer.setPageMargins(QMarginsF(15, 15, 15, 15), QPageLayout.Unit.Millimeter)

                doc = QTextDocument()
                doc.setHtml(html)
                doc.print(pdf_printer)
                print(f"[PRINT] PDF temporal creado: {tmp_pdf}")
            except Exception:
                try:
                    os.remove(tmp_pdf)
                except Exception:
                    pass
                raise

            images = render_pdf_to_qimages(tmp_pdf, dpi=300)
            if not images:
                raise RuntimeError("No se pudo rasterizar el PDF para imprimir.")

            def is_virtual(info: QPrinterInfo) -> bool:
                name = (info.printerName() or "").upper()
                return any(k in name for k in ("PDF", "XPS", "ONENOTE", "FAX"))

            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            chosen: Optional[QPrinterInfo] = None
            try:
                chosen = QPrinterInfo.defaultPrinter()
            except Exception:
                chosen = None

            if chosen and chosen.printerName() and not is_virtual(chosen):
                printer.setPrinterName(chosen.printerName())
                print(f"[PRINT] Usando impresora: {chosen.printerName()}")
            else:
                dlg = QPrintDialog(printer, self)
                dlg.setWindowTitle("Imprimir (como imagen)")
                if dlg.exec() != QDialog.DialogCode.Accepted:
                    try:
                        os.remove(tmp_pdf)
                    except Exception:
                        pass
                    return

            printer.setResolution(300)
            painter = QPainter()
            if not painter.begin(printer):
                raise RuntimeError("No se pudo iniciar la impresión.")

            try:
                for idx, image in enumerate(images):
                    target_rect = painter.viewport()
                    scaled_size = image.size()
                    scaled_size.scale(target_rect.size(), Qt.AspectRatioMode.KeepAspectRatio)

                    painter.setViewport(
                        target_rect.x(),
                        target_rect.y(),
                        scaled_size.width(),
                        scaled_size.height()
                    )
                    painter.setWindow(image.rect())
                    painter.drawImage(0, 0, image)

                    if idx < len(images) - 1:
                        printer.newPage()
            finally:
                painter.end()

            QMessageBox.information(self, "Imprimir", "Documento enviado a la impresora como imagen (300 ppp).")
            print("[PRINT] Trabajo enviado correctamente (rasterizado 300 ppp).")
        except Exception as e:
            QMessageBox.critical(self, "Imprimir", f"No se pudo imprimir:\n{human_ex(e)}")
        finally:
            try:
                if 'tmp_pdf' in locals() and os.path.exists(tmp_pdf):
                    os.remove(tmp_pdf)
            except Exception:
                pass

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
        self._populate_cameras()
        self.cam_combo.currentIndexChanged.connect(self._on_cam_changed)

        self.btn_rot_l = QPushButton("↺"); self.btn_rot_l.setFixedWidth(30)
        self.btn_rot_r = QPushButton("↻"); self.btn_rot_r.setFixedWidth(30)
        self.btn_rot_l.clicked.connect(lambda: self._set_rot(-90))
        self.btn_rot_r.clicked.connect(lambda: self._set_rot(+90))

        self.chk_autocrop = QCheckBox("Auto-crop")
        self.chk_autocrop.setChecked(True)

        self.btn_capture = QPushButton("Capturar imagen")
        self.btn_capture.clicked.connect(self._capture_frame)

        cam_row.addWidget(self.cam_combo, 3)
        cam_row.addSpacing(6)
        cam_row.addWidget(self.btn_rot_l); cam_row.addWidget(self.btn_rot_r)
        cam_row.addSpacing(8)
        cam_row.addWidget(self.chk_autocrop)
        cam_row.addSpacing(8)
        cam_row.addWidget(self.btn_capture)
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
        self.body_edit.setAcceptRichText(True)
        self.body_edit.setPlaceholderText("Se generará automáticamente con IA al cargar una imagen.")
        right.addWidget(self.body_edit, 1)

        actions = QHBoxLayout()
        self.btn_pdf = QPushButton("DESCARGAR PDF")
        self.btn_print = QPushButton("IMPRIMIR")
        self.btn_send = QPushButton("ENVIAR")
        self.btn_pdf.clicked.connect(self.export_pdf)
        self.btn_print.clicked.connect(self.print_preview)
        self.btn_send.clicked.connect(self.send_email)
        actions.addStretch(1)
        actions.addWidget(self.btn_pdf)
        actions.addWidget(self.btn_print)
        actions.addWidget(self.btn_send)
        right.addLayout(actions)

        cols.addLayout(right, 1)

        self._regen_timer = QTimer(self); self._regen_timer.setSingleShot(True)
        self._regen_timer.timeout.connect(lambda: self._maybe_autogenerate(trigger="name-change"))
        self.to_edit.textChanged.connect(lambda: self._regen_timer.start(500))

        self.drop.image_loaded.connect(lambda _p: self._maybe_autogenerate(trigger="image"))
        self.body_edit.textChanged.connect(self._sync_raw_from_editor)

    def _sync_raw_from_editor(self):
        self._body_raw = self.body_edit.toPlainText().strip()

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
        self._stop_preview()
        if cam_index is None or int(cam_index) == -1:
            self.drop.reset_placeholder()
        else:
            self._start_preview(int(cam_index))

    def _start_preview(self, index: int):
        self._cap = cv2.VideoCapture(index, cv2.CAP_DSHOW)
        if not self._cap.isOpened():
            QMessageBox.critical(self, "Cámara", "No se pudo abrir la cámara seleccionada.")
            self._cap = None
            return
        self._cam_timer.start()

    def _stop_preview(self):
        self._cam_timer.stop()
        if self._cap:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None

    def _apply_rot_crop(self, frame):
        if self._rot == 90:
            frame = cv2.rotate(frame, cv2.ROTATE_90_CLOCKWISE)
        elif self._rot == 180:
            frame = cv2.rotate(frame, cv2.ROTATE_180)
        elif self._rot == 270:
            frame = cv2.rotate(frame, cv2.ROTATE_90_COUNTERCLOCKWISE)

        if self.chk_autocrop.isChecked() and self.drop.width() > 0 and self.drop.height() > 0:
            h, w = frame.shape[:2]
            target_aspect = self.drop.width() / max(1, self.drop.height())
            frame_aspect = w / max(1, h)
            if abs(frame_aspect - target_aspect) > 1e-3:
                if frame_aspect > target_aspect:
                    new_w = int(h * target_aspect)
                    x0 = max(0, (w - new_w) // 2)
                    frame = frame[:, x0:x0 + new_w]
                else:
                    new_h = int(w / target_aspect)
                    y0 = max(0, (h - new_h) // 2)
                    frame = frame[y0:y0 + new_h, :]
        return frame

    def _render_preview(self, frame):
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb.shape
        qimg = QImage(rgb.data, w, h, ch * w, QImage.Format.Format_RGB888)
        self.drop.show_qimage(qimg)

    def _update_preview(self):
        if not self._cap:
            return
        ret, frame = self._cap.read()
        if not ret:
            return
        self._last_frame = frame
        frame = self._apply_rot_crop(self._last_frame.copy())
        self._render_preview(frame)

    def _set_rot(self, delta: int):
        self._rot = (self._rot + delta) % 360
        if self._last_frame is not None:
            frame = self._apply_rot_crop(self._last_frame.copy())
            self._render_preview(frame)

    def _capture_frame(self):
        cam_index = self.cam_combo.currentData()
        if cam_index is None or cam_index == -1:
            QMessageBox.warning(self, "Cámara", "Selecciona una cámara distinta de 'Ninguna'.")
            return
        if self._last_frame is None:
            QMessageBox.warning(self, "Cámara", "Aún no hay imagen de la cámara.")
            return
        tmp = resource_path("_captura_tmp.png")
        frame = self._apply_rot_crop(self._last_frame.copy())
        cv2.imwrite(tmp, frame)
        self.drop.load_path(tmp)
        try:
            os.remove(tmp)
        except Exception:
            pass

    # IA
    def _img_and_mime_or_none(self):
        return (self.drop.bytes, self.drop.mime) if self.drop.bytes else (None, None)

    def _maybe_autogenerate(self, trigger: str):
        api_key = self.main.settings.value("accounts/google_api_key", "", str)
        print(f"[IA-Presupuestos] Trigger: {trigger} · API {'sí' if api_key else 'no'}")
        if not api_key:
            return

        if trigger.startswith("name") and (self._body_raw or self.body_edit.toPlainText().strip()):
            name_now = self.to_edit.text().strip()
            if name_now:
                self._body_raw = (self._body_raw or self.body_edit.toPlainText()).replace("[Nombre del destinatario]", name_now)
                self._render_body_from_raw()
                return

        img, mime = self._img_and_mime_or_none()
        if not img:
            print("[IA-Presupuestos] No hay imagen para procesar.")
            return

        name_now = self.to_edit.text().strip()
        self.subject_edit.setPlaceholderText("Generando asunto con IA…")
        self.body_edit.setPlaceholderText("Generando cuerpo con IA…")
        self.btn_pdf.setEnabled(False); self.btn_send.setEnabled(False); self.btn_print.setEnabled(False)

        self._ai_worker = EmailGenWorker(api_key, name_now, img, mime)
        self._ai_worker.finished.connect(self._on_ai_ready)
        self._ai_worker.failed.connect(self._on_ai_failed)
        self._ai_worker.start()
        print("[IA-Presupuestos] Worker iniciado.")

    def _apply_owner_name(self, text: str) -> str:
        owner = self.main.settings.value("accounts/owner_name", "", str).strip()
        if not owner:
            return text
        placeholders = (
            "[Tu Nombre]", "[Tu nombre]", "[Tu Nombre/Empresa]", "[Nombre/Empresa]",
            "[Nombre de Empresa]", "[Nombre de empresa]", "[nombre de empresa]",
            "[Empresa]", "[empresa]", "[Mi Empresa]", "[mi empresa]"
        )
        for ph in placeholders:
            text = text.replace(ph, owner)
        return text

    def update_owner_name(self, owner: str):
        owner = (owner or "").strip()
        if not owner:
            return
        placeholders = (
            "[Tu Nombre]", "[Tu nombre]", "[Tu Nombre/Empresa]", "[Nombre/Empresa]",
            "[Nombre de Empresa]", "[Nombre de empresa]", "[nombre de empresa]",
            "[Empresa]", "[empresa]", "[Mi Empresa]", "[mi empresa]"
        )
        changed = False
        if self._body_raw:
            new_raw = self._body_raw
            for ph in placeholders:
                new_raw = new_raw.replace(ph, owner)
            if new_raw != self._body_raw:
                self._body_raw = new_raw
                changed = True
        else:
            current = self.body_edit.toPlainText()
            new = current
            for ph in placeholders:
                new = new.replace(ph, owner)
            if new != current:
                self._body_raw = new
                changed = True
        if changed:
            self._render_body_from_raw()

    def _render_body_from_raw(self):
        raw = (self._body_raw or "").strip()
        html_div = self._body_text_to_html(raw)
        self.body_edit.blockSignals(True)
        self.body_edit.setHtml(html_div)
        self.body_edit.blockSignals(False)

    def _on_ai_ready(self, subject: str, body: str):
        print("[IA-Presupuestos] Resultado recibido.")
        self.subject_edit.blockSignals(True)
        if not self.subject_edit.text().strip():
            self.subject_edit.setText(subject)
        name_now = self.to_edit.text().strip()
        if name_now:
            body = body.replace("[Nombre del destinatario]", name_now)
        self._body_raw = self._apply_owner_name(body)
        self.subject_edit.blockSignals(False)

        self._render_body_from_raw()

        self.btn_pdf.setEnabled(True); self.btn_send.setEnabled(True); self.btn_print.setEnabled(True)

    def _on_ai_failed(self, msg: str):
        print("[IA-Presupuestos] Error:", msg)
        self.btn_pdf.setEnabled(True); self.btn_send.setEnabled(True); self.btn_print.setEnabled(True)

    def _body_text_to_html(self, body_text: str) -> str:
        rendered = markdown_table_to_html(body_text)
        if rendered is None:
            safe = html_escape(body_text)
            safe = "".join(
                f"<p style='margin:0 0 10px'>{line}</p>" if line.strip()
                else "<p style='margin:0 0 12px'>&nbsp;</p>"
                for line in safe.split("\n")
            )
            content = safe
        else:
            content = rendered

        return (
            "<div style='font-family:Arial,Helvetica,sans-serif;"
            "font-size:12pt;color:#222;line-height:1.6'>"
            f"{content}"
            "</div>"
        )

    def export_pdf(self):
        default_dir = self.main.settings.value("paths/pdf_dir", os.getcwd(), str) or os.getcwd()
        default_path = os.path.join(default_dir, "presupuesto.pdf")
        path, _ = QFileDialog.getSaveFileName(self, "Guardar PDF", default_path, "PDF (*.pdf)")
        if not path:
            return
        if not path.lower().endswith(".pdf"):
            path += ".pdf"
        try:
            subj = (self.subject_edit.text().strip() or "Solicitud de cotización")
            html = self._compose_html_document(subj, self._body_raw or self.body_edit.toPlainText())

            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(path)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            printer.setPageMargins(QMarginsF(15, 15, 15, 15), QPageLayout.Unit.Millimeter)

            doc = QTextDocument()
            doc.setHtml(html)
            doc.print(printer)
            print(f"[PDF] Guardado en: {path}")
        except Exception as e:
            QMessageBox.critical(self, "PDF", f"No se pudo generar el PDF:\n{human_ex(e)}")

    def print_document(self):
        try:
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setPageMargins(QMarginsF(52, 52, 52, 52))
            subj = (self.subject_edit.text().strip() or "Solicitud de cotización")
            body_text = self._apply_owner_name(self.body_edit.toPlainText().strip())
            body_html = self._body_text_to_html(body_text)
            html = f"""
            <html><head><meta charset='utf-8'>
            <style>
                body{{font-family:Arial,Helvetica,sans-serif; font-size:12pt; color:#222; line-height:1.6}}
                h2{{margin:0 0 8px}}
                hr{{border:none; border-top:1px solid #ddd; margin:0 0 14px}}
            </style>
            </head><body>
            <h2>{html_escape(subj)}</h2>
            <hr>
            {body_html}
            </body></html>
            """
            doc = QTextDocument()
            doc.setHtml(html)
            doc.print(printer)
            QMessageBox.information(self, "Imprimir", "Documento enviado a la impresora predeterminada.")
        except Exception as e:
            QMessageBox.critical(self, "Imprimir", f"No se pudo imprimir:\n{human_ex(e)}")

    def send_email(self):
        server = self.main.settings.value("accounts/smtp_server", "smtp.gmail.com", str)
        port = int(self.main.settings.value("accounts/smtp_port", 587, int))
        user = self.main.settings.value("accounts/smtp_email", "", str)
        password = self.main.settings.value("accounts/smtp_password", "", str)
        to_addr = self.mail_edit.text().strip()
        subject = self.subject_edit.text().strip()
        body_text = self._apply_owner_name((self._body_raw or self.body_edit.toPlainText()).strip())

        if not (server and port and user and password and to_addr and subject and body_text):
            QMessageBox.warning(self, "Correo", "Completa credenciales SMTP en Configuración y los campos del correo.")
            return

        try:
            from email.message import EmailMessage
            msg = EmailMessage()
            msg["From"] = user
            msg["To"] = to_addr
            msg["Subject"] = subject

            msg.set_content(body_text)
            html_body = self._body_text_to_html(body_text)
            msg.add_alternative(f"""<html><body>{html_body}</body></html>""", subtype="html")

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

    def stop_camera(self):
        self._stop_preview()

# ------------------- Configuración -------------------

def _bytes_to_gb(b: int) -> int:
    return int(round(b / (1024**3)))

def get_system_info() -> str:
    cpu = platform.processor() or "Desconocido"
    ram_txt = "Desconocida"
    if psutil:
        try:
            ram_txt = f"Capacidad Total: {_bytes_to_gb(psutil.virtual_memory().total)} GB"
        except Exception:
            pass

    gpus = []
    try:
        if sys.platform.startswith("win"):
            out = subprocess.run(
                ["wmic", "path", "win32_VideoController", "get", "Name"],
                capture_output=True, text=True
            ).stdout
            for line in out.splitlines():
                line = line.strip()
                if line and line.upper() != "NAME":
                    gpus.append(line)
    except Exception:
        pass
    if not gpus:
        gpus.append("No detectado")

    cuda_lines = ["No se detectó una GPU NVIDIA..."]
    try:
        nsmi = subprocess.run(["nvidia-smi", "--query-gpu=cuda_version", "--format=csv,noheader"],
                              capture_output=True, text=True)
        if nsmi.returncode == 0 and nsmi.stdout.strip():
            ver = nsmi.stdout.strip().splitlines()[0].trim() if hasattr(str, "trim") else nsmi.stdout.strip().splitlines()[0].strip()
            cuda_lines = [ "Se detectó una GPU NVIDIA...", f"Versión de CUDA instalada: {ver}" ]
    except Exception:
        pass

    block = []
    block.append("[PROCESADOR]")
    block.append(cpu)
    block.append("")
    block.append("[MEMORIA RAM]")
    block.append(ram_txt)
    block.append("")
    block.append("[TARJETA GRÁFICA (GPU)]")
    block.extend(gpus)
    block.append("")
    block.append("[VERSIÓN DE CUDA]")
    block.extend(cuda_lines)

    return "\n".join(block)

class SettingsDialog(QDialog):
    def __init__(self, parent_main):
        super().__init__(parent_main)
        self.main = parent_main
        self.setWindowTitle("Configuración")
        self.setMinimumSize(780, 600)
        lay = QVBoxLayout(self)

        # Topbar
        top = QHBoxLayout()
        back = QPushButton("←"); back.setFixedSize(36, 28); back.clicked.connect(self.accept)
        title = QLabel("Configuración")
        f = QFont(); f.setPointSize(14); title.setFont(f)
        top.addWidget(back); top.addStretch(1); top.addWidget(title); top.addStretch(5)
        lay.addLayout(top)

        self.tabs = QTabWidget()
        lay.addWidget(self.tabs, 1)

        self._tab_general()
        self._tab_accounts_api()

        QTimer.singleShot(400, self._auto_validate_all)

    def _tab_general(self):
        w = QWidget(); l = QVBoxLayout(w)

        # ---- Ruta TGP (EXE/LNK)
        grp1 = QGroupBox("Ruta actual de TGP (exe/lnk)")
        g1 = QVBoxLayout(grp1)
        current_tgp = self.main.settings.value("paths/tgp_path", "", str)
        self.ed_prog = QLineEdit(current_tgp)
        self.ed_prog.setReadOnly(True)
        b1 = QPushButton("Cambiar…")

        def choose_prog():
            start_dir = os.path.dirname(self.ed_prog.text()) if self.ed_prog.text() else os.getcwd()
            path, _ = QFileDialog.getOpenFileName(
                self, "Selecciona ejecutable o acceso directo de TGP",
                start_dir,
                "Ejecutables y accesos directos (*.exe *.lnk);;Todos los archivos (*.*)"
            )
            if path:
                self.ed_prog.setText(path)
                self.main.settings.setValue("paths/tgp_path", path)

        b1.clicked.connect(choose_prog)
        g1.addWidget(self.ed_prog); g1.addWidget(b1)

        # ---- Carpeta PDF
        grp2 = QGroupBox("Ruta actual para guardar los PDF")
        g2 = QVBoxLayout(grp2)
        self.ed_pdf = QLineEdit(self.main.settings.value("paths/pdf_dir", os.getcwd(), str))
        self.ed_pdf.setReadOnly(True)
        b2 = QPushButton("Elegir carpeta de destino")
        def choose_pdf_dir():
            d = QFileDialog.getExistingDirectory(self, "Selecciona carpeta PDF", self.ed_pdf.text() or os.getcwd())
            if d:
                self.ed_pdf.setText(d)
                self.main.settings.setValue("paths/pdf_dir", d)
        b2.clicked.connect(choose_pdf_dir)
        g2.addWidget(self.ed_pdf); g2.addWidget(b2)

        # ---- Sistema / Versión
        grp_sys = QGroupBox("Características:")
        gl = QVBoxLayout(grp_sys)
        self.lbl_specs = QLabel(get_system_info())
        self.lbl_specs.setStyleSheet(
            "font-family:Consolas, monospace; background:#f6f7fb; "
            "border:1px solid #e1e3ea; border-radius:8px; padding:10px;"
        )
        self.lbl_specs.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        gl.addWidget(self.lbl_specs)

        ver_row = QVBoxLayout()
        self.lbl_version = QLabel(f"Versión: v{APP_VERSION}")
        self.btn_update = QPushButton("Comprobando actualizaciones…")
        self.btn_update.setEnabled(False)
        self.btn_update.setMinimumHeight(36)
        self.btn_update.setStyleSheet(
            "QPushButton{background:#eee; border:1px solid #d6d9e0; border-radius:6px}"
            "QPushButton:disabled{color:#999;}"
        )
        self.btn_update.clicked.connect(self._start_update)
        ver_row.addWidget(self.lbl_version)
        ver_row.addWidget(self.btn_update)

        l.addWidget(grp1); l.addWidget(grp2); l.addWidget(grp_sys); l.addLayout(ver_row); l.addStretch(1)
        self.tabs.addTab(w, "General")

        self._update_info = None
        self._upd = UpdateCheckerWorker()
        self._upd.finished.connect(self._on_update_check)
        self._upd.start()

    def _tab_accounts_api(self):
        w = QWidget(); l = QVBoxLayout(w)

        grp_owner = QGroupBox("Identidad (firma)")
        fo = QFormLayout(grp_owner)
        self.ed_owner = QLineEdit(self.main.settings.value("accounts/owner_name", "", str))
        fo.addRow("Nombre o Empresa:", self.ed_owner)

        grp_g = QGroupBox("Google API (Gemini)")
        fg = QFormLayout(grp_g)
        self.ed_api = QLineEdit(self.main.settings.value("accounts/google_api_key", "", str))
        self.lbl_api_state = QLabel("●"); self.lbl_api_state.setStyleSheet("color: grey; font-size:18px;")
        self.ed_api.setEchoMode(QLineEdit.EchoMode.Password)
        fg.addRow("API Key:", self.ed_api)
        fg.addRow("Estado:", self.lbl_api_state)

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

        l.addWidget(grp_owner)
        l.addWidget(grp_g)
        l.addWidget(grp_s)
        l.addStretch(1)
        self.tabs.addTab(w, "Cuentas / API")

        # Timers automáticos
        def _commit_owner():
            name = self.ed_owner.text().strip()
            self.main.settings.setValue("accounts/owner_name", name)
            try:
                self.main.page_b.update_owner_name(name)
            except Exception:
                pass

        self._t_owner = QTimer(self); self._t_owner.setSingleShot(True)
        self._t_owner.timeout.connect(_commit_owner)
        self.ed_owner.textChanged.connect(lambda _=None: self._t_owner.start(400))

        self._t_api = QTimer(self); self._t_api.setSingleShot(True)
        self._t_api.timeout.connect(self._validate_api_auto)
        self.ed_api.textChanged.connect(lambda _=None: self._t_api.start(700))

        self._t_smtp = QTimer(self); self._t_smtp.setSingleShot(True)
        self._t_smtp.timeout.connect(self._validate_smtp_auto)
        for wdg in (self.ed_smtp_server, self.ed_smtp_port, self.ed_smtp_email, self.ed_smtp_pass):
            wdg.textChanged.connect(lambda _=None: self._t_smtp.start(800))

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
            self.btn_update.setText(f"Actualizar a v{tag}")
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

        top = QHBoxLayout()
        top.addSpacing(36)
        title = QLabel("BALADA PACKAGING .S.L.U.")
        f = QFont(); f.setPointSize(18); f.setBold(True); title.setFont(f)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        top.addWidget(title, 1)
        gear = QPushButton("⚙️"); gear.setFixedSize(36, 28)
        gear.clicked.connect(self.open_settings.emit)
        top.addWidget(gear)
        lay.addLayout(top)

        imgs = QHBoxLayout()
        self.img_a = ClickLabel(); self._load_img(self.img_a, "multimedia/Opcion_A.png")
        self.img_b = ClickLabel(); self._load_img(self.img_b, "multimedia/Opcion_B.png")

        self.img_a.clicked.connect(self.go_a.emit)
        self.img_b.clicked.connect(self.go_b.emit)

        imgs.addWidget(self.img_a, 1)
        imgs.addWidget(self.img_b, 1)
        lay.addLayout(imgs, 1)

        foot = QHBoxLayout()
        la = QLabel("TG Profesional: Ingresar datos al sistema"); la.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lb = QLabel("Solicitud de cotizacion a proveedores"); lb.setAlignment(Qt.AlignmentFlag.AlignCenter)
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

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

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

    def closeEvent(self, e):
        try:
            self.page_b.stop_camera()
        except Exception:
            pass
        try:
            self.page_a.stop_camera()
        except Exception:
            pass
        super().closeEvent(e)

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
