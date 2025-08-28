# main.py
# Python 3.10 · Windows · PyQt6
# App base para gestionar el lanzamiento de TGProfesional (.lnk),
# detectar si ya está abierta y sentar las bases de automatización (ventanas/teclado/clicks).

import os
import sys
import time
import subprocess
from typing import Optional, Tuple, List

from PyQt6 import QtWidgets, QtCore

# --- Dependencias externas para Windows / automatización ---
# psutil para detectar procesos en ejecución
import psutil

# pywin32 para resolver .lnk (atajo de Windows)
try:
    import win32com.client  # pywin32
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

# pywinauto para automatización de ventanas (UIA)
try:
    from pywinauto import Application, Desktop
    from pywinauto.keyboard import send_keys as pwa_send_keys
    HAS_PYWINAUTO = True
except Exception:
    HAS_PYWINAUTO = False

# pyautogui para clicks/teclado a nivel de pantalla (fallback genérico)
try:
    import pyautogui
    HAS_PYAUTOGUI = True
except Exception:
    HAS_PYAUTOGUI = False


def resolve_windows_shortcut(lnk_path: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Resuelve un .lnk a (target_exe, arguments, working_dir).
    Requiere pywin32. Si no está disponible o falla, retorna (None, None, None).
    """
    if not HAS_WIN32:
        return None, None, None
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(lnk_path)
        target = shortcut.Targetpath or None
        args = shortcut.Arguments or None
        work = shortcut.WorkingDirectory or None
        return target, args, work
    except Exception:
        return None, None, None


def any_tg_window_present(window_hints: List[str]) -> bool:
    """
    Usa pywinauto (si disponible) para buscar ventanas cuyo título coincida
    con alguno de los patrones proporcionados.
    """
    if not HAS_PYWINAUTO:
        return False
    try:
        desk = Desktop(backend="uia")
        for pattern in window_hints:
            wins = desk.windows(title_re=pattern, visible_only=True)
            if wins:
                return True
    except Exception:
        pass
    return False


def is_process_running_for_target(target_exe: Optional[str], name_hints: List[str]) -> bool:
    """
    Verifica si el proceso objetivo ya se está ejecutando.
    - Si se conoce target_exe: intenta comparación por ruta/nombre.
    - Además, usa 'name_hints' (p. ej., "TGProfesional.exe") como respaldo.
    Maneja permisos/errores con cuidado.
    """
    target_basename = os.path.basename(target_exe).lower() if target_exe else None
    hints_lower = [h.lower() for h in name_hints if h]

    for proc in psutil.process_iter(["name", "exe", "cmdline"]):
        try:
            name = (proc.info.get("name") or "").lower()
            exe = (proc.info.get("exe") or "")
            cmd = proc.info.get("cmdline") or []

            # Coincidencia por ruta exacta del ejecutable
            if target_exe and exe:
                try:
                    # Evita excepciones con samefile si alguna ruta no existe/acceso denegado
                    if os.path.exists(exe) and os.path.exists(target_exe):
                        if os.path.samefile(exe, target_exe):
                            return True
                except Exception:
                    pass

            # Coincidencia por nombre base
            if target_basename and name and (name == target_basename):
                return True

            # Coincidencia con hints (nombre aproximado)
            if name and any(name == hint for hint in hints_lower):
                return True

            if cmd:
                first = os.path.basename(cmd[0]).lower()
                if target_basename and first == target_basename:
                    return True
                if any(first == hint for hint in hints_lower):
                    return True

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue

    return False


class AutomationBridge:
    """
    Capa de automatización para usar luego:
    - Traer ventanas al frente
    - Enviar teclas
    - Clicks (con pywinauto o pyautogui)
    """
    def __init__(self, window_title_patterns: Optional[List[str]] = None):
        self.window_title_patterns = window_title_patterns or [r"TG\s*Profesional", r"TGProfesional", r"TGP"]

    def find_window(self):
        if not HAS_PYWINAUTO:
            return None
        try:
            desk = Desktop(backend="uia")
            for pattern in self.window_title_patterns:
                wins = desk.windows(title_re=pattern, visible_only=True)
                if wins:
                    return wins[0]
        except Exception:
            return None
        return None

    def bring_to_front(self) -> bool:
        w = self.find_window()
        if w is None:
            return False
        try:
            w.set_focus()
            return True
        except Exception:
            return False

    def send_keys(self, keys: str) -> bool:
        """
        Envía teclas a la ventana activa.
        - Con pywinauto: pwa_send_keys("^a{ENTER}") por ejemplo
        - Con pyautogui: pyautogui.write / pyautogui.hotkey
        """
        if HAS_PYWINAUTO:
            try:
                pwa_send_keys(keys)
                return True
            except Exception:
                pass

        if HAS_PYAUTOGUI:
            try:
                # En pyautogui no existe la misma sintaxis, esto es un fallback simplificado.
                pyautogui.typewrite(keys)
                return True
            except Exception:
                pass
        return False

    def click_center(self) -> bool:
        """
        Ejemplo simple: click en el centro de la pantalla (para demostrar que pyautogui está disponible).
        Lo refinaremos luego con coordenadas/controles.
        """
        if not HAS_PYAUTOGUI:
            return False
        try:
            w, h = pyautogui.size()
            pyautogui.click(w // 2, h // 2)
            return True
        except Exception:
            return False


class MainWindow(QtWidgets.QMainWindow):
    # 1) Variable con la ruta .lnk (puedes cambiarla si es necesario)
    APP_SHORTCUT = r"C:\TGProfesional\TGProfesional.lnk"

    # Hints para ventana/proceso (ajustables si cambian los títulos/nombres)
    WINDOW_TITLE_PATTERNS = [r"TG\s*Profesional", r"TGProfesional", r"TGP"]
    PROCESS_NAME_HINTS = ["TGProfesional.exe", "TGP.exe"]

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestor TGProfesional")
        self.resize(420, 160)

        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)

        self.btn = QtWidgets.QPushButton("Procesar", central)
        self.btn.setMinimumHeight(40)
        self.btn.clicked.connect(self.on_procesar)

        self.status = QtWidgets.QLabel("Listo.", central)
        self.status.setWordWrap(True)

        layout = QtWidgets.QVBoxLayout(central)
        layout.addWidget(self.btn)
        layout.addWidget(self.status)

        # Capa de automatización (para futuro uso)
        self.automator = AutomationBridge(self.WINDOW_TITLE_PATTERNS)

    # 3) Lógica al pulsar "Procesar"
    def on_procesar(self):
        lnk_path = self.APP_SHORTCUT

        # Intentar resolver el .lnk a su .exe real
        target_exe, target_args, working_dir = resolve_windows_shortcut(lnk_path) if os.path.exists(lnk_path) else (None, None, None)

        # Comprobar si ya está abierto
        already_running = False

        # a) Comprobación por proceso (si sabemos el .exe), con hints de nombre
        if target_exe:
            already_running = is_process_running_for_target(target_exe, self.PROCESS_NAME_HINTS)
        else:
            # Si no pudimos resolver el .lnk, igual intentamos por hints de nombre
            already_running = is_process_running_for_target(None, self.PROCESS_NAME_HINTS)

        # b) Comprobación por ventana (si pywinauto está disponible)
        if not already_running:
            if any_tg_window_present(self.WINDOW_TITLE_PATTERNS):
                already_running = True

        if already_running:
            QtWidgets.QMessageBox.information(self, "Estado", "La aplicación ya estaba abierta.")
            self.status.setText("Detectado: la aplicación ya está abierta.")
            # Opcional: traerla al frente (no obligatorio)
            self.automator.bring_to_front()
            return

        # Si no está abierta → informamos y la abrimos
        QtWidgets.QMessageBox.information(self, "Estado", "La aplicación no estaba abierta. Intentaré abrirla ahora.")

        opened = self.launch_app(lnk_path, target_exe, target_args, working_dir)

        if opened:
            self.status.setText("Se ha enviado la orden para abrir la aplicación.")
            # No bloqueamos la UI, pero hacemos una comprobación breve tras 2 s
            QtCore.QTimer.singleShot(2000, self.post_launch_check)
        else:
            self.status.setText("No se pudo abrir la aplicación.")
            QtWidgets.QMessageBox.critical(self, "Error", "No se pudo abrir la aplicación. Revisa la ruta del acceso directo o permisos.")

    def post_launch_check(self):
        """Verificación ligera tras el intento de apertura, para feedback al usuario."""
        # Intento detectar ventana o proceso tras el lanzamiento
        lnk_path = self.APP_SHORTCUT
        target_exe, _, _ = resolve_windows_shortcut(lnk_path) if os.path.exists(lnk_path) else (None, None, None)

        running = False
        if target_exe:
            running = is_process_running_for_target(target_exe, self.PROCESS_NAME_HINTS)
        else:
            running = is_process_running_for_target(None, self.PROCESS_NAME_HINTS)

        if not running and any_tg_window_present(self.WINDOW_TITLE_PATTERNS):
            running = True

        if running:
            QtWidgets.QMessageBox.information(self, "Estado", "Aplicación abierta correctamente.")
            self.automator.bring_to_front()
        else:
            QtWidgets.QMessageBox.warning(self, "Aviso", "No se detectó la aplicación tras el intento de apertura.")

    def launch_app(self, lnk_path: str, target_exe: Optional[str], target_args: Optional[str], working_dir: Optional[str]) -> bool:
        """
        Lanza la app:
        - Preferimos abrir el .lnk con os.startfile (Windows abre el destino).
        - Si falla, intentamos lanzar el .exe resuelto con argumentos.
        """
        try:
            if os.path.exists(lnk_path):
                os.startfile(lnk_path)  # ShellExecute por detrás
                return True

            # Fallback si no existe el .lnk pero sí el .exe
            if target_exe and os.path.exists(target_exe):
                args = []
                if target_args:
                    # División simple por espacios; si hay comillas/espacios complejos, ajústalo a tus necesidades
                    args = target_args.split()
                creationflags = subprocess.CREATE_NEW_PROCESS_GROUP if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP") else 0
                subprocess.Popen([target_exe, *args], cwd=working_dir or os.path.dirname(target_exe) or None,
                                 creationflags=creationflags, close_fds=False)
                return True

        except Exception:
            pass

        return False


def main():
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
