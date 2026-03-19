import customtkinter as ctk
import tkintermapview
import webbrowser
import ctypes
import os
import sys
import time
import threading
import copy
import logging
import calendar as pycalendar
import queue
import subprocess
import re
from datetime import datetime, timedelta, timezone
from PIL import Image, ImageTk, ImageDraw
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog
from urllib.parse import quote
from uuid import uuid4
from pathlib import Path

from services.storage_employees import load_employees, save_employees
from services.storage_vehicles import (
    delete_trailer,
    delete_vehicle,
    load_vehicles,
    save_vehicles,
    upsert_trailer,
    upsert_vehicle,
)
from services.storage_tours import (
    filter_tours_by_date,
    filter_tours_by_range,
    format_date,
    load_tours,
    parse_date,
    save_tours,
    tour_assignment_count,
)
from services.routing_service import estimate_distance_km, get_travel_segment
from services.schedule_planner import compute_schedule
from services.time_utils import format_time, parse_time, validate_time_window
from services.geocoding_service import GeocodingService
from services.json_storage import InvalidJsonFileError, atomic_write_json, load_json_file
from services.map_route_service import RouteServiceError, fetch_route_path
from services.route_optimization_service import optimize_stop_order
from services.pin_storage import load_pins as load_pin_records, save_pins as save_pin_records
from services.sql_import_service import DEFAULT_SQL_DATA_DIR, fetch_order_rows, infer_database_name_from_data_dir
from backup_manager import BackupManager
from config.update_config import APP_NAME, SHOW_UPDATE_PAGE_IN_MENU
from pages.vehicles_page import VehiclesPage
from pages.settings_page import SettingsPage
from pages.update_page import UpdatePage
from services.version_service import get_runtime_update_context, set_update_log_dir
from settings_manager import SettingsManager

try:
    from tkcalendar import Calendar as TkCalendar
except Exception:
    TkCalendar = None

try:
    from customtkinter.windows.widgets.ctk_scrollable_frame import CTkScrollableFrame as _CTkScrollableFrame
except Exception:
    _CTkScrollableFrame = None

# ============================================================
# Modern UI / Theme
# ============================================================
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


logger = logging.getLogger(__name__)
GPS_TRACKING_URL = "https://map.ktrac.ch/"


def _patch_scrollable_frame_mousewheel_guard():
    if _CTkScrollableFrame is None:
        return

    original = getattr(_CTkScrollableFrame, "check_if_master_is_canvas", None)
    if not callable(original) or getattr(original, "_gawela_safe_patch", False):
        return

    def _safe_check_if_master_is_canvas(self, widget):
        candidate = widget
        if isinstance(candidate, str):
            try:
                candidate = self.nametowidget(candidate)
            except Exception:
                return False
        if not hasattr(candidate, "master"):
            return False
        return original(self, candidate)

    _safe_check_if_master_is_canvas._gawela_safe_patch = True
    _CTkScrollableFrame.check_if_master_is_canvas = _safe_check_if_master_is_canvas


_patch_scrollable_frame_mousewheel_guard()


def _get_runtime_dirs() -> tuple[str, str]:
    bundle_dir = os.path.dirname(os.path.abspath(__file__))
    if getattr(sys, "frozen", False):
        bundle_dir = os.path.dirname(sys.executable)
        appdata_root = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        user_data_dir = os.path.join(appdata_root, "GAWELA", "Tourenplaner")
        return bundle_dir, user_data_dir
    return bundle_dir, bundle_dir


def _get_bundle_dir_candidates(bundle_dir: str) -> list[str]:
    candidates = []
    for value in (
        bundle_dir,
        os.path.dirname(os.path.abspath(__file__)),
        os.path.dirname(getattr(sys, "executable", "") or ""),
        getattr(sys, "_MEIPASS", ""),
        os.getcwd(),
    ):
        path = os.path.abspath(value) if value else ""
        if not path or path in candidates:
            continue
        candidates.append(path)
        parent = os.path.dirname(path)
        if parent and parent not in candidates:
            candidates.append(parent)
    return candidates


def _resolve_runtime_asset_path(bundle_dir: str, *relative_parts: str) -> str:
    for candidate in _get_bundle_dir_candidates(bundle_dir):
        asset_path = os.path.join(candidate, *relative_parts)
        if os.path.exists(asset_path):
            return asset_path
    return os.path.join(bundle_dir, *relative_parts)


def _find_webview2_runtime_in_root(root_dir: str) -> str | None:
    if not root_dir:
        return None
    try:
        root_dir = os.path.abspath(root_dir)
    except Exception:
        return None
    if not os.path.isdir(root_dir):
        return None

    direct_executable = os.path.join(root_dir, "msedgewebview2.exe")
    if os.path.exists(direct_executable):
        return root_dir

    child_dirs = []
    try:
        child_dirs = [
            entry.path
            for entry in os.scandir(root_dir)
            if entry.is_dir()
        ]
    except Exception:
        return None

    def _rank(path: str):
        name = os.path.basename(path)
        numeric_parts = []
        for part in name.split("."):
            numeric_parts.append(int(part) if part.isdigit() else -1)
        return tuple(numeric_parts), name.lower()

    for candidate in sorted(child_dirs, key=_rank, reverse=True):
        executable = os.path.join(candidate, "msedgewebview2.exe")
        if os.path.exists(executable):
            return candidate
    return None


def _resolve_system_webview2_runtime_path() -> str | None:
    standard_roots = [
        r"C:\Program Files (x86)\Microsoft\EdgeWebView\Application",
        r"C:\Program Files\Microsoft\EdgeWebView\Application",
    ]
    for root in standard_roots:
        runtime_dir = _find_webview2_runtime_in_root(root)
        if runtime_dir:
            return runtime_dir
    return None


def _resolve_webview2_runtime_path(bundle_dir: str) -> str | None:
    candidate_roots = [
        _resolve_runtime_asset_path(bundle_dir, "assets", "webview2"),
        _resolve_runtime_asset_path(bundle_dir, "webview2"),
        _resolve_runtime_asset_path(bundle_dir, "assets", "WebView2"),
    ]
    for candidate in candidate_roots:
        runtime_dir = _find_webview2_runtime_in_root(candidate)
        if runtime_dir:
            return runtime_dir
    return _resolve_system_webview2_runtime_path()


def _apply_webview2_runtime_environment(bundle_dir: str, env: dict | None = None) -> tuple[dict, str | None]:
    target_env = dict(os.environ if env is None else env)
    runtime_path = _resolve_webview2_runtime_path(bundle_dir)
    if runtime_path:
        target_env["WEBVIEW2_BROWSER_EXECUTABLE_FOLDER"] = runtime_path
    return target_env, runtime_path


def _is_runtime_inside_bundle(bundle_dir: str, runtime_path: str | None) -> bool:
    if not runtime_path:
        return False
    try:
        runtime_abs = os.path.abspath(runtime_path)
    except Exception:
        return False

    bundled_roots = [
        _resolve_runtime_asset_path(bundle_dir, "assets", "webview2"),
        _resolve_runtime_asset_path(bundle_dir, "webview2"),
        _resolve_runtime_asset_path(bundle_dir, "assets", "WebView2"),
    ]
    for root in bundled_roots:
        try:
            root_abs = os.path.abspath(root)
            if os.path.commonpath([runtime_abs, root_abs]) == root_abs:
                return True
        except Exception:
            continue
    return False


def _run_gps_webview_window(url: str):
    import webview

    bundle_dir, config_dir = _get_runtime_dirs()
    storage_path = os.path.join(config_dir, "gps-webview-profile")
    os.makedirs(storage_path, exist_ok=True)

    _, runtime_path = _apply_webview2_runtime_environment(bundle_dir)
    if runtime_path:
        webview.settings["WEBVIEW2_RUNTIME_PATH"] = runtime_path

    webview.settings["OPEN_EXTERNAL_LINKS_IN_BROWSER"] = True
    webview.create_window(
        "GPS",
        url=url,
        text_select=True,
        zoomable=True,
        confirm_close=False,
        width=1600,
        height=960,
    )
    webview.start(
        gui="edgechromium",
        private_mode=False,
        storage_path=storage_path,
    )


def _attach_embedded_webview_window(child_hwnd: int, parent_hwnd: int, width: int, height: int):
    user32 = ctypes.windll.user32
    gwl_style = -16
    ws_child = 0x40000000
    ws_visible = 0x10000000
    ws_popup = 0x80000000
    sw_show = 5

    get_style = user32.GetWindowLongPtrW
    set_style = user32.SetWindowLongPtrW
    get_style.restype = ctypes.c_longlong
    set_style.restype = ctypes.c_longlong

    style = int(get_style(child_hwnd, gwl_style))
    style = (style | ws_child | ws_visible) & ~ws_popup
    set_style(child_hwnd, gwl_style, style)
    user32.SetParent(child_hwnd, parent_hwnd)
    user32.MoveWindow(child_hwnd, 0, 0, int(width), int(height), True)
    user32.ShowWindow(child_hwnd, sw_show)


def _get_parent_client_size(parent_hwnd: int) -> tuple[int, int]:
    class RECT(ctypes.Structure):
        _fields_ = [
            ("left", ctypes.c_long),
            ("top", ctypes.c_long),
            ("right", ctypes.c_long),
            ("bottom", ctypes.c_long),
        ]

    rect = RECT()
    user32 = ctypes.windll.user32
    if not user32.GetClientRect(int(parent_hwnd), ctypes.byref(rect)):
        return 400, 320
    width = max(int(rect.right - rect.left), 400)
    height = max(int(rect.bottom - rect.top), 320)
    return width, height


def _run_gps_embedded_webview_host(parent_hwnd: int, width: int, height: int, url: str):
    import pythoncom
    import clr
    from webview import _state as webview_state
    from webview import settings as webview_settings
    from webview.platforms.edgechromium import EdgeChrome
    from webview.window import Window

    clr.AddReference("System.Windows.Forms")
    clr.AddReference("System.Drawing")

    import System.Windows.Forms as WinForms
    from System.Drawing import Point, Size

    bundle_dir, config_dir = _get_runtime_dirs()
    storage_path = os.path.join(config_dir, "gps-webview-profile")
    os.makedirs(storage_path, exist_ok=True)

    _, runtime_path = _apply_webview2_runtime_environment(bundle_dir)
    if runtime_path:
        webview_settings["WEBVIEW2_RUNTIME_PATH"] = runtime_path

    pythoncom.CoInitialize()
    try:
        webview_state["private_mode"] = False
        webview_state["debug"] = False
        webview_state["user_agent"] = None
        webview_state["ssl"] = False

        form = WinForms.Form()
        form.Text = "GPS"
        form.FormBorderStyle = getattr(WinForms.FormBorderStyle, "None")
        form.ShowInTaskbar = False
        form.StartPosition = WinForms.FormStartPosition.Manual
        form.Location = Point(0, 0)
        form.Size = Size(max(int(width), 400), max(int(height), 320))

        handle = int(form.Handle.ToInt64())
        _attach_embedded_webview_window(handle, int(parent_hwnd), width, height)

        window = Window(str(uuid4()), "GPS", url, width=form.Width, height=form.Height)
        window.real_url = url
        browser = EdgeChrome(form, window, storage_path)

        timer = WinForms.Timer()
        timer.Interval = 150
        last_size = {"value": (0, 0)}

        def _resize(*_args):
            try:
                target_width, target_height = _get_parent_client_size(int(parent_hwnd))
                current_size = (int(target_width), int(target_height))
                if current_size == last_size["value"]:
                    return
                last_size["value"] = current_size
                _attach_embedded_webview_window(handle, int(parent_hwnd), target_width, target_height)
            except Exception:
                pass

        timer.Tick += _resize
        timer.Start()

        def _on_closed(*_args):
            try:
                timer.Stop()
            except Exception:
                pass
            try:
                browser.webview.Dispose()
            except Exception:
                pass

        form.FormClosed += _on_closed
        form.Show()
        _resize()
        WinForms.Application.Run(form)
    finally:
        pythoncom.CoUninitialize()


def _run_auxiliary_mode_from_argv() -> bool:
    args = list(sys.argv[1:])
    if not args:
        return False

    if args[0] == "--gps-embed":
        if len(args) < 4:
            raise SystemExit("Missing arguments for --gps-embed")
        parent_hwnd = int(str(args[1]).strip())
        width = int(str(args[2]).strip())
        height = int(str(args[3]).strip())
        url = args[4].strip() if len(args) > 4 and str(args[4]).strip() else GPS_TRACKING_URL
        _run_gps_embedded_webview_host(parent_hwnd, width, height, url)
        return True

    if args[0] != "--gps-webview":
        return False

    url = args[1].strip() if len(args) > 1 and str(args[1]).strip() else GPS_TRACKING_URL
    try:
        _run_gps_webview_window(url)
    except Exception as exc:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("GPS", f"Das native GPS-Fenster konnte nicht gestartet werden:\n{exc}")
        try:
            root.destroy()
        except Exception:
            pass
    return True


def _terminate_process_gracefully(process: subprocess.Popen | None, timeout: float = 5.0) -> bool:
    if process is None:
        return True
    try:
        if process.poll() is not None:
            return True
    except Exception:
        return True

    try:
        process.terminate()
        process.wait(timeout=timeout)
        return True
    except Exception:
        try:
            process.kill()
            process.wait(timeout=max(timeout, 1.0))
            return True
        except Exception:
            return False


class Theme:
    BG = ("#F5F7FB", "#0F1115")
    PANEL = ("#FFFFFF", "#151822")
    PANEL_2 = ("#F0F2F8", "#11131B")
    BORDER = ("#E3E7F2", "#262B3A")
    OVERLAY_PANEL = ("#FFFFFF", "#1A1D26")
    OVERLAY_PANEL_2 = ("#F8FAFC", "#151922")
    OVERLAY_BORDER = ("#E5E7EB", "#2A2F3A")
    SELECTION = ("#EAF2FF", "#1F2A3A")
    CONFLICT_BG = ("#FEE2E2", "#3A1F24")
    TEXT = ("#0F172A", "#E7EAF3")
    SUBTEXT = ("#475569", "#A9B0C3")
    ACCENT = ("#800080", "#800080")
    ACCENT_HOVER = ("#660066", "#990099")
    SUCCESS = ("#16A34A", "#22C55E")
    SUCCESS_HOVER = ("#15803D", "#16A34A")
    DANGER = ("#DC2626", "#EF4444")
    DANGER_HOVER = ("#B91C1C", "#DC2626")
    WARNING = ("#D97706", "#F59E0B")
    WARNING_HOVER = ("#B45309", "#D97706")
    MUTED_BTN = ("#334155", "#2B2F3A")
    MUTED_BTN_HOVER = ("#1F2937", "#3A4050")
    SCROLLBAR_TRACK = "transparent"
    SCROLLBAR_BUTTON = SUBTEXT
    SCROLLBAR_HOVER = ACCENT

    @staticmethod
    def resolve(color):
        if not isinstance(color, tuple):
            return color
        return color[1] if ctk.get_appearance_mode() == "Dark" else color[0]


def _font(size: int = 14, weight: str = "normal"):
    return ctk.CTkFont(size=size, weight=weight)


def _scrollable_frame_kwargs():
    return {
        "scrollbar_fg_color": Theme.SCROLLBAR_TRACK,
        "scrollbar_button_color": Theme.SCROLLBAR_BUTTON,
        "scrollbar_button_hover_color": Theme.SCROLLBAR_HOVER,
    }


def _scrollbar_kwargs():
    return {
        "fg_color": Theme.SCROLLBAR_TRACK,
        "button_color": Theme.SCROLLBAR_BUTTON,
        "button_hover_color": Theme.SCROLLBAR_HOVER,
        "corner_radius": 999,
    }


class EmbeddedWebView2Frame(tk.Frame):
    def __init__(self, master, app, *, url: str, status_callback=None):
        super().__init__(master, bg=Theme.resolve(Theme.PANEL_2), highlightthickness=0, bd=0)
        self.app = app
        self.url = str(url or GPS_TRACKING_URL).strip() or GPS_TRACKING_URL
        self.status_callback = status_callback
        self._started = False
        self._disposed = False
        self._ready = False
        self._helper_process = None
        self._ui_queue = queue.Queue()
        self._storage_path = os.path.join(self.app.config_dir, "gps-webview-profile")
        os.makedirs(self._storage_path, exist_ok=True)

        self.placeholder = ctk.CTkLabel(
            self,
            text="WebView2 wird initialisiert...",
            font=_font(14, "bold"),
            text_color=Theme.SUBTEXT,
        )
        self.placeholder.place(relx=0.5, rely=0.5, anchor="center")

        self.bind("<Map>", self._on_map, add="+")
        self.bind("<Configure>", self._on_configure, add="+")
        self.bind("<Destroy>", self._on_destroy, add="+")
        self.after(50, self._drain_ui_queue)
        self.after(400, self._poll_helper_process)

    def _set_status(self, message: str):
        if callable(self.status_callback):
            try:
                self.status_callback(str(message or ""))
            except Exception:
                pass

    def _queue_ui_call(self, callback):
        try:
            self._ui_queue.put_nowait(callback)
        except Exception:
            pass

    def _drain_ui_queue(self):
        if self._disposed:
            return
        while True:
            try:
                callback = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            try:
                callback()
            except Exception:
                logger.exception("Queued GPS UI callback failed.")
        self.after(50, self._drain_ui_queue)

    def _show_placeholder(self, message: str):
        try:
            self.placeholder.configure(text=message)
            self.placeholder.place(relx=0.5, rely=0.5, anchor="center")
        except Exception:
            pass

    def _hide_placeholder(self):
        try:
            self.placeholder.place_forget()
        except Exception:
            pass

    def _on_map(self, _event=None):
        self._ensure_started()

    def _on_destroy(self, event=None):
        if event is not None and getattr(event, "widget", None) is not self:
            return
        self.dispose()

    def _on_configure(self, _event=None):
        return

    def _ensure_started(self):
        if self._started or self._disposed:
            return
        if os.name != "nt":
            self._handle_runtime_error("Die eingebettete GPS-Ansicht wird nur unter Windows unterstuetzt.")
            return

        try:
            parent_hwnd = int(self.winfo_id())
        except Exception:
            self.after(120, self._ensure_started)
            return
        if not parent_hwnd:
            self.after(120, self._ensure_started)
            return

        self._started = True
        self._show_placeholder("WebView2 wird initialisiert...")
        self._set_status("WebView2 wird gestartet...")
        self._start_helper_process(parent_hwnd)

    def _start_helper_process(self, parent_hwnd: int):
        width = max(self.winfo_width(), 400)
        height = max(self.winfo_height(), 320)
        try:
            command = self.app.get_gps_embed_helper_command(parent_hwnd, width, height, self.url)
            startupinfo = None
            creationflags = 0
            if os.name == "nt":
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

            env, _runtime_path = _apply_webview2_runtime_environment(self.app.base_dir)
            self._helper_process = subprocess.Popen(
                command,
                cwd=self.app.base_dir,
                env=env,
                startupinfo=startupinfo,
                creationflags=creationflags,
            )
            self.after(700, self._handle_webview_ready)
        except Exception as exc:
            logger.exception("Embedded WebView2 helper could not be started.")
            self._handle_runtime_error(str(exc))

    def _handle_webview_ready(self):
        if self._disposed:
            return
        self._ready = True
        self._hide_placeholder()
        runtime_path = _resolve_webview2_runtime_path(self.app.base_dir)
        if runtime_path and _is_runtime_inside_bundle(self.app.base_dir, runtime_path):
            self._set_status(f"Eingebettete WebView2 aktiv. Gebundene Runtime: {runtime_path}")
        elif runtime_path:
            self._set_status(f"Eingebettete WebView2 aktiv. Systemruntime: {runtime_path}")
        else:
            self._set_status("Eingebettete WebView2 aktiv.")

    def _handle_webview_closed(self):
        if self._disposed:
            return
        self._ready = False
        self._show_placeholder("Die eingebettete GPS-Ansicht wurde geschlossen.")
        self._set_status("Die eingebettete GPS-Ansicht wurde geschlossen.")

    def _handle_runtime_error(self, message: str):
        self._started = False
        self._ready = False
        self._show_placeholder(f"WebView2 konnte nicht gestartet werden:\n{message}")
        self._set_status(f"WebView2-Fehler: {message}")

    def _handle_navigation_starting(self, sender, args):
        try:
            target = str(args.Uri)
        except Exception:
            target = self.url
        self._queue_ui_call(lambda: self._set_status(f"Lade: {target}"))

    def _handle_navigation_completed(self, sender, args):
        return

    def _poll_helper_process(self):
        if self._disposed:
            return
        process = self._helper_process
        if process is not None:
            try:
                code = process.poll()
            except Exception:
                code = None
            if code is not None:
                self._helper_process = None
                if not self._disposed:
                    self._ready = False
                    self._started = False
                    self._show_placeholder("WebView2 konnte nicht gestartet werden oder wurde beendet.")
                    self._set_status(f"WebView2-Hostprozess beendet (Code {code}).")
        self.after(400, self._poll_helper_process)

    def navigate(self, url: str):
        target = str(url or "").strip() or self.url
        previous_url = self.url
        self.url = target
        if not self._ready or self._helper_process is None or self._helper_process.poll() is not None:
            self._ensure_started()
            return
        if target == previous_url:
            return
        self.reload()

    def reload(self):
        if self._helper_process is not None:
            _terminate_process_gracefully(self._helper_process, timeout=2.0)
            self._helper_process = None
        self._ready = False
        self._started = False
        self._show_placeholder("WebView2 wird initialisiert...")
        if self._disposed:
            return
        if not self.winfo_ismapped():
            return
        self._ensure_started()

    def dispose(self):
        if self._disposed:
            return
        self._disposed = True
        if self._helper_process is not None:
            _terminate_process_gracefully(self._helper_process, timeout=2.0)
            self._helper_process = None


def _normalize_date_string(value) -> str:
    parsed = parse_date(value)
    return format_date(parsed) if parsed else ""


def _display_date_string(value) -> str:
    parsed = parse_date(value)
    if not parsed:
        return ""
    return parsed.strftime("%d.%m.%Y")


def _date_for_calendar(year: int, month: int, day: int) -> str:
    try:
        return format_date(datetime(year, month, day).date())
    except Exception:
        return ""


def _shift_month(year: int, month: int, delta: int) -> tuple[int, int]:
    total = ((int(year) * 12) + (int(month) - 1)) + int(delta)
    shifted_year, shifted_month_index = divmod(total, 12)
    return shifted_year, shifted_month_index + 1


GERMAN_MONTH_NAMES = [
    "",
    "Januar",
    "Februar",
    "März",
    "April",
    "Mai",
    "Juni",
    "Juli",
    "August",
    "September",
    "Oktober",
    "November",
    "Dezember",
]


SERVICE_MINUTE_OPTIONS = ["0", "5", "10", "15", "20", "30", "45", "60", "90", "120"]
TIME_HOUR_OPTIONS = [f"{hour:02d}" for hour in range(24)]
TIME_MINUTE_OPTIONS = [f"{minute:02d}" for minute in range(0, 60, 5)]
DEFAULT_QUICK_ACCESS_ITEMS = ["action:export_route", "action:import_sql", "", ""]
DELIVERY_TYPE_OPTIONS = ["Frei Bordsteinkante", "Mit Verteilung", "mit Verteilung & montage"]
DEFAULT_DELIVERY_TYPE = DELIVERY_TYPE_OPTIONS[0]
_DELIVERY_TYPE_CANONICAL = {
    "frei bordsteinkante": "Frei Bordsteinkante",
    "mit verteilung": "Mit Verteilung",
    "mit verteilung & montage": "mit Verteilung & montage",
}
_NON_MAP_DELIVERY_TYPE_CANONICAL = {
    "selbstabholung": "Selbstabholung",
    "post": "Post",
    "fracht mit spediteur": "Fracht mit Spediteur",
    "direktlieferung frei bordsteinkante": "Direktlieferung frei Bordsteinkante",
    "direktlieferung ohne warenverteilung": "Direktlieferung ohne Warenverteilung",
    "fracht-tresor-bordstein": "Fracht-Tresor-Bordstein",
    "fracht-tresor-verwendung": "Fracht-Tresor-Verwendung",
}
_SQL_DELIVERY_CODE_TO_LABEL = {
    "fracht_o_vert": "Frei Bordsteinkante",
    "fracht_m_vert": "Mit Verteilung",
    "fracht_m_vert_mont": "mit Verteilung & montage",
    "selbstabholung": "Selbstabholung",
    "post": "Post",
    "fracht_mit_spediteur": "Fracht mit Spediteur",
    "fracht-tresor-bordstein": "Direktlieferung frei Bordsteinkante",
    "fracht-tresor-verwendung": "Direktlieferung ohne Warenverteilung",
}
_SQL_NON_MAP_DELIVERY_CODES = {
    "selbstabholung",
    "post",
    "fracht_mit_spediteur",
    "fracht-tresor-bordstein",
    "fracht-tresor-verwendung",
}
NON_MAP_ORDER_FILTER_OPTIONS = [
    "Post",
    "Spediteur",
    "Selbstabholung",
    "Fracht-Tresor-Bordstein",
    "Fracht-Tresor-Verwendung",
]
COUNTRY_HINT_BY_CODE = {
    "CH": "Schweiz",
    "DE": "Deutschland",
    "AT": "Oesterreich",
    "PL": "Polen",
    "CZ": "Tschechien",
    "ES": "Spanien",
}
PRODUCT_QTY_PREFIX_RE = re.compile(
    r"^\s*(\d+)\s*(?:x|stk\.?|stueck|stück)?\s*[:\-]?\s*(.+)$",
    flags=re.IGNORECASE,
)


def _normalize_delivery_type(value) -> str:
    text = str(value or "").strip()
    if not text:
        return DEFAULT_DELIVERY_TYPE
    normalized = _DELIVERY_TYPE_CANONICAL.get(text.casefold())
    if normalized:
        return normalized
    non_map = _NON_MAP_DELIVERY_TYPE_CANONICAL.get(text.casefold())
    if non_map:
        return non_map
    return DEFAULT_DELIVERY_TYPE


def _country_hint_from_code(code: str) -> str:
    text = str(code or "").strip().upper()
    if not text:
        return "Schweiz"
    return COUNTRY_HINT_BY_CODE.get(text, text)


def _resolve_sql_delivery_handling(delivery_code: str) -> tuple[str, bool]:
    code = str(delivery_code or "").strip().casefold()
    if not code:
        return DEFAULT_DELIVERY_TYPE, True
    label = _SQL_DELIVERY_CODE_TO_LABEL.get(code, DEFAULT_DELIVERY_TYPE)
    include_on_map = code not in _SQL_NON_MAP_DELIVERY_CODES
    return label, include_on_map


def _is_non_map_delivery_type(value) -> bool:
    key = str(value or "").strip().casefold()
    if not key:
        return False
    return key in _SQL_NON_MAP_DELIVERY_CODES or key in _NON_MAP_DELIVERY_TYPE_CANONICAL


def _normalize_non_map_order_category(value) -> str:
    key = str(value or "").strip().casefold()
    if not key:
        return ""
    if "post" in key:
        return "Post"
    if "selbstabholung" in key:
        return "Selbstabholung"
    if "spediteur" in key or "fracht_mit_spediteur" in key:
        return "Spediteur"
    if "tresor-bordstein" in key or "direktlieferung frei bordsteinkante" in key:
        return "Fracht-Tresor-Bordstein"
    if "tresor-verwendung" in key or "direktlieferung ohne warenverteilung" in key:
        return "Fracht-Tresor-Verwendung"
    return ""


def _first_non_empty_text(*values) -> str:
    for value in values:
        text = str(value or "").strip()
        if text:
            return text
    return ""


def _normalize_order_and_delivery_addresses(data: dict | None) -> dict:
    payload = dict(data or {})

    delivery_name = _first_non_empty_text(payload.get("LieferadresseName"), payload.get("Name"))
    delivery_street = _first_non_empty_text(payload.get("LieferadresseStrasse"), payload.get("Strasse"))
    delivery_plz = _first_non_empty_text(payload.get("LieferadressePLZ"), payload.get("PLZ"))
    delivery_ort = _first_non_empty_text(payload.get("LieferadresseOrt"), payload.get("Ort"))

    order_name = _first_non_empty_text(
        payload.get("AuftragsadresseName"),
        payload.get("AuftragName"),
    )
    order_plz = _first_non_empty_text(
        payload.get("AuftragsadressePLZ"),
        payload.get("AuftragPLZ"),
    )
    order_ort = _first_non_empty_text(
        payload.get("AuftragsadresseOrt"),
        payload.get("AuftragOrt"),
    )
    if not (order_name or order_plz or order_ort):
        order_name = delivery_name
        order_plz = delivery_plz
        order_ort = delivery_ort

    payload["AuftragsadresseName"] = order_name
    payload["AuftragsadressePLZ"] = order_plz
    payload["AuftragsadresseOrt"] = order_ort

    payload["LieferadresseName"] = delivery_name
    payload["LieferadresseStrasse"] = delivery_street
    payload["LieferadressePLZ"] = delivery_plz
    payload["LieferadresseOrt"] = delivery_ort

    # Compatibility fields used in existing UI and routing logic.
    payload["Name"] = delivery_name
    payload["Strasse"] = delivery_street
    payload["PLZ"] = delivery_plz
    payload["Ort"] = delivery_ort
    return payload


def _format_order_address(data: dict | None) -> str:
    payload = _normalize_order_and_delivery_addresses(data)
    name = str(payload.get("AuftragsadresseName") or "").strip()
    city_line = " ".join(
        value for value in [str(payload.get("AuftragsadressePLZ") or "").strip(), str(payload.get("AuftragsadresseOrt") or "").strip()] if value
    ).strip()
    if name and city_line:
        return f"{name}, {city_line}"
    return name or city_line


def _format_delivery_address(data: dict | None) -> str:
    payload = _normalize_order_and_delivery_addresses(data)
    name = str(payload.get("LieferadresseName") or "").strip()
    street = str(payload.get("LieferadresseStrasse") or "").strip()
    city_line = " ".join(
        value for value in [str(payload.get("LieferadressePLZ") or "").strip(), str(payload.get("LieferadresseOrt") or "").strip()] if value
    ).strip()
    line2 = ", ".join(value for value in [street, city_line] if value).strip()
    if name and line2:
        return f"{name}, {line2}"
    return name or line2


def _split_delivery_name_lines(name: str) -> list[str]:
    text = str(name or "").strip()
    if not text:
        return []
    for separator in (" / ", " | ", ";"):
        if separator in text:
            return [part.strip() for part in text.split(separator) if part.strip()]
    if "," in text:
        left, right = [part.strip() for part in text.split(",", 1)]
        parts = [part for part in (left, right) if part]
        if len(parts) >= 2:
            return parts
    return [text]


def _format_delivery_address_multiline(data: dict | None) -> str:
    payload = _normalize_order_and_delivery_addresses(data)
    name = str(payload.get("LieferadresseName") or "").strip()
    street = str(payload.get("LieferadresseStrasse") or "").strip()
    city_line = " ".join(
        value for value in [str(payload.get("LieferadressePLZ") or "").strip(), str(payload.get("LieferadresseOrt") or "").strip()] if value
    ).strip()
    lines = []
    lines.extend(_split_delivery_name_lines(name))
    if street:
        lines.append(street)
    if city_line:
        lines.append(city_line)
    return "\n".join(lines)


def _set_text_input_value(widget, value: str):
    text = str(value or "")
    setter = getattr(widget, "set", None)
    if callable(setter):
        setter(text)
        return
    try:
        widget.delete(0, "end")
        widget.insert(0, text)
    except Exception:
        pass


class TimeDropdownField(ctk.CTkFrame):
    def __init__(self, master, *, options: list[str], width: int = 76, height: int = 36):
        super().__init__(master, fg_color="transparent")
        self.options = [str(option) for option in options]
        self._callbacks = []
        self._popup = None
        self._outside_click_binding_id = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)

        self.shell = ctk.CTkFrame(
            self,
            corner_radius=14,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        self.shell.grid(row=0, column=0, sticky="ew")
        self.shell.grid_columnconfigure(0, weight=1)
        self.shell.grid_columnconfigure(1, weight=0)

        self.entry = ctk.CTkEntry(
            self.shell,
            height=height,
            width=width,
            corner_radius=0,
            border_width=0,
            fg_color=Theme.PANEL,
            text_color=Theme.TEXT,
            font=_font(12, "bold"),
            justify="center",
        )
        self.entry.grid(row=0, column=0, sticky="ew", padx=(10, 4), pady=4)
        self.entry.bind("<KeyRelease>", lambda _event: self._emit_change(), add="+")
        self.entry.bind("<Down>", self._open_popup_event, add="+")

        self.button = ctk.CTkButton(
            self.shell,
            text="▾",
            width=28,
            height=height - 4,
            corner_radius=12,
            fg_color=Theme.PANEL,
            hover_color=Theme.PANEL_2,
            text_color=Theme.ACCENT,
            font=_font(12, "bold"),
            command=self.toggle_popup,
        )
        self.button.grid(row=0, column=1, padx=4, pady=2)

    def get(self) -> str:
        return str(self.entry.get() or "").strip()

    def set(self, value: str):
        self.entry.delete(0, "end")
        self.entry.insert(0, str(value or "").strip())

    def bind_change(self, callback):
        if callable(callback):
            self._callbacks.append(callback)

    def _emit_change(self):
        for callback in list(self._callbacks):
            callback()

    def _open_popup_event(self, _event=None):
        self.open_popup()
        return "break"

    def toggle_popup(self):
        if self._popup is not None and self._popup.winfo_exists():
            self.close_popup()
        else:
            self.open_popup()

    def open_popup(self):
        if self._popup is not None and self._popup.winfo_exists():
            return

        popup = tk.Toplevel(self)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        transparent_mask = "#010203"
        popup.configure(bg=transparent_mask)
        try:
            popup.wm_attributes("-transparentcolor", transparent_mask)
        except tk.TclError:
            popup.configure(bg=Theme.resolve(Theme.BG))
        self._popup = popup

        self.update_idletasks()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 4
        width = max(self.winfo_width(), 92)
        popup.geometry(f"{width}x220+{x}+{y}")

        shell = ctk.CTkFrame(
            popup,
            corner_radius=14,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        shell.pack(fill="both", expand=True)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(
            shell,
            corner_radius=12,
            fg_color=Theme.PANEL,
            **_scrollable_frame_kwargs(),
        )
        scroll.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        scroll.grid_columnconfigure(0, weight=1)

        current = self.get()
        for option in self.options:
            selected = option == current
            ctk.CTkButton(
                scroll,
                text=option,
                height=32,
                corner_radius=10,
                fg_color=Theme.SELECTION if selected else "transparent",
                hover_color=Theme.PANEL_2,
                text_color=Theme.TEXT,
                anchor="w",
                command=lambda value=option: self.select_option(value),
            ).grid(sticky="ew", padx=4, pady=3)

        popup.bind("<Escape>", lambda _event: self.close_popup(), add="+")
        popup.bind("<FocusOut>", lambda _event: self.after(60, self._close_if_focus_left), add="+")
        root = self.winfo_toplevel()
        self._outside_click_binding_id = root.bind("<ButtonPress-1>", self._handle_click_outside, add="+")
        popup.focus_force()

    def _close_if_focus_left(self):
        popup = self._popup
        if popup is None or not popup.winfo_exists():
            return
        focus_widget = popup.focus_displayof()
        if focus_widget is None:
            self.close_popup()

    def close_popup(self):
        popup = self._popup
        self._popup = None
        root = self.winfo_toplevel()
        if self._outside_click_binding_id:
            try:
                root.unbind("<ButtonPress-1>", self._outside_click_binding_id)
            except Exception:
                pass
            self._outside_click_binding_id = None
        if popup is None:
            return
        try:
            if popup.winfo_exists():
                popup.destroy()
        except tk.TclError:
            pass

    def select_option(self, value: str):
        self.set(value)
        self._emit_change()
        self.close_popup()

    def _handle_click_outside(self, event):
        popup = self._popup
        if popup is None or not popup.winfo_exists():
            return

        clicked_widget = event.widget
        if self._is_descendant(clicked_widget, popup) or self._is_descendant(clicked_widget, self):
            return
        self.close_popup()

    def _is_descendant(self, widget, ancestor) -> bool:
        current = widget
        while current is not None:
            if current == ancestor:
                return True
            current = getattr(current, "master", None)
        return False


class TimeInput(ctk.CTkFrame):
    def __init__(self, master, *, width: int | None = None, height: int = 36):
        super().__init__(master, fg_color="transparent")
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=0)
        self.grid_columnconfigure(2, weight=1)

        field_width = 76 if width is None else max(64, int(width // 2) - 8)
        self.hour_field = TimeDropdownField(self, options=TIME_HOUR_OPTIONS, width=field_width, height=height)
        self.hour_field.grid(row=0, column=0, sticky="ew")

        self.separator = ctk.CTkLabel(self, text=":", font=_font(14, "bold"), text_color=Theme.TEXT)
        self.separator.grid(row=0, column=1, padx=6)

        self.minute_field = TimeDropdownField(self, options=TIME_MINUTE_OPTIONS, width=field_width, height=height)
        self.minute_field.grid(row=0, column=2, sticky="ew")

        self.set("08:00")

    def get(self) -> str:
        hour = self.hour_field.get()
        minute = self.minute_field.get()
        if not hour and not minute:
            return ""
        return f"{hour}:{minute}"

    def set(self, value: str):
        parsed = parse_time(value)
        if parsed is None:
            text = str(value or "").strip()
            if ":" in text:
                hour, minute = (part.strip() for part in text.split(":", 1))
            else:
                hour, minute = text, ""
        else:
            hour = f"{parsed.hour:02d}"
            minute = f"{parsed.minute:02d}"
        self.hour_field.set(hour)
        self.minute_field.set(minute)

    def bind_change(self, callback):
        self.hour_field.bind_change(callback)
        self.minute_field.bind_change(callback)


# ============================================================
# UI Components
# ============================================================
class NavButton(ctk.CTkButton):
    def __init__(self, master, text, command, selected=False, compact_icon=None, compact_image=None):
        self.full_text = text
        self.compact_icon = compact_icon or text[:1]
        self.compact_image = compact_image
        super().__init__(
            master,
            text=text,
            height=42,
            corner_radius=12,
            anchor="w",
            command=command,
            font=_font(14, "bold" if selected else "normal"),
            fg_color=Theme.ACCENT if selected else "transparent",
            hover_color=Theme.ACCENT_HOVER if selected else Theme.PANEL_2,
            text_color=("white", "white") if selected else Theme.TEXT,
            border_width=1,
            border_color=Theme.BORDER if not selected else Theme.ACCENT,
        )
        self._compact = False

    def set_selected(self, selected: bool):
        self.configure(
            font=_font(14, "bold" if selected else "normal"),
            fg_color=Theme.ACCENT if selected else "transparent",
            hover_color=Theme.ACCENT_HOVER if selected else Theme.PANEL_2,
            text_color=("white", "white") if selected else Theme.TEXT,
            border_color=Theme.BORDER if not selected else Theme.ACCENT,
        )

    def set_compact(self, compact: bool):
        compact = bool(compact)
        if self._compact == compact:
            return
        self._compact = compact
        self.configure(
            text="" if compact and self.compact_image is not None else (self.compact_icon if compact else self.full_text),
            image=self.compact_image if compact else None,
            anchor="center" if compact else "w",
            width=40 if compact else 0,
            height=38 if compact else 42,
            corner_radius=10 if compact else 12,
        )


class DropdownSection(ctk.CTkFrame):
    """Modern Accordion Block"""

    def __init__(self, master, title: str, default_open: bool = False):
        super().__init__(
            master,
            corner_radius=14,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        self._open = False

        self.header = ctk.CTkButton(
            self,
            text=f"▸ {title}",
            anchor="w",
            height=40,
            corner_radius=12,
            font=_font(14, "bold"),
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.toggle,
        )
        self.header.pack(fill="x", padx=10, pady=10)

        self.body = ctk.CTkFrame(self, corner_radius=12, fg_color="transparent")




    def toggle(self):
        self._open = not self._open
        if self._open:
            self.header.configure(text=self.header.cget("text").replace("▸", "▾"))
            self.body.pack(fill="x", padx=10, pady=(0, 10))
        else:
            self.header.configure(text=self.header.cget("text").replace("▾", "▸"))
            self.body.pack_forget()


# ============================================================
# Calendar UI
# ============================================================
class TourCalendarWidget(ctk.CTkFrame):
    """Fallback-Kalender im App-Theme.

    Markierungslogik:
    - Orange: genau 1 geplanter Mitarbeitereinsatz an diesem Datum.
    - Rot: 2 oder mehr geplante Mitarbeitereinsätze an diesem Datum.
    Legacy-Touren ohne `employee_ids` werden als 1 Einsatz gezählt, damit bestehende Daten sichtbar bleiben.
    """

    def __init__(self, master, app, on_date_selected=None, on_date_activated=None):
        super().__init__(
            master,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        self.app = app
        self.on_date_selected = on_date_selected
        self.on_date_activated = on_date_activated
        today = datetime.now()
        self.display_year = today.year
        self.display_month = today.month
        self.selected_date = None
        self._day_payload = {}
        self._last_click_date = None
        self._last_click_ts = 0.0
        self._day_buttons = {}
        self._visible_month_keys = None
        self._payload_revision = -1

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, padx=12, pady=(12, 8), sticky="ew")
        header.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            header,
            text="<",
            width=36,
            height=32,
            corner_radius=12,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self._prev_month,
        ).grid(row=0, column=0, sticky="w")

        self.month_label = ctk.CTkLabel(header, text="", font=_font(14, "bold"), text_color=Theme.TEXT)
        self.month_label.grid(row=0, column=1, sticky="n")

        ctk.CTkButton(
            header,
            text=">",
            width=36,
            height=32,
            corner_radius=12,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self._next_month,
        ).grid(row=0, column=2, sticky="e")

        self.months_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.months_frame.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")
        self.months_frame.grid_columnconfigure((0, 1), weight=1, uniform="calendar-month")
        self.months_frame.grid_rowconfigure(0, weight=1)

        self.month_cards = []
        for col in range(2):
            month_card = ctk.CTkFrame(
                self.months_frame,
                corner_radius=14,
                fg_color=Theme.PANEL,
                border_width=1,
                border_color=Theme.BORDER,
            )
            month_card.grid(row=0, column=col, padx=(0, 6) if col == 0 else (6, 0), sticky="nsew")
            month_card.grid_columnconfigure(0, weight=1)
            month_card.grid_rowconfigure(2, weight=1)

            month_title = ctk.CTkLabel(month_card, text="", font=_font(13, "bold"), text_color=Theme.TEXT)
            month_title.grid(row=0, column=0, padx=12, pady=(12, 8), sticky="w")

            week = ctk.CTkFrame(month_card, fg_color="transparent")
            week.grid(row=1, column=0, padx=10, pady=(0, 4), sticky="ew")
            for i in range(7):
                week.grid_columnconfigure(i, weight=1)

            for week_col, label in enumerate(["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]):
                ctk.CTkLabel(week, text=label, font=_font(12, "bold"), text_color=Theme.SUBTEXT).grid(
                    row=0, column=week_col, padx=2, pady=(0, 4), sticky="ew"
                )

            days_frame = ctk.CTkFrame(month_card, fg_color="transparent")
            days_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")
            for i in range(7):
                days_frame.grid_columnconfigure(i, weight=1)
            for i in range(6):
                days_frame.grid_rowconfigure(i, weight=1)

            self.month_cards.append((month_title, days_frame))

        self.refresh()

    def _prev_month(self):
        self.display_year, self.display_month = _shift_month(self.display_year, self.display_month, -1)
        self.refresh()

    def _next_month(self):
        self.display_year, self.display_month = _shift_month(self.display_year, self.display_month, 1)
        self.refresh()

    def _get_payload_map(self):
        return self.app.get_calendar_payload_map()

    def _pick_colors(self, date_key: str):
        payload = self._day_payload.get(date_key)
        if not payload:
            return Theme.PANEL, Theme.BORDER, Theme.TEXT
        if payload["assignments"] >= 2:
            return Theme.DANGER, Theme.DANGER_HOVER, ("#FFFFFF", "#FFFFFF")
        return Theme.WARNING, Theme.WARNING_HOVER, ("#FFFFFF", "#FFFFFF")

    def _visible_month_tuple(self) -> tuple[tuple[int, int], tuple[int, int]]:
        return (
            (self.display_year, self.display_month),
            _shift_month(self.display_year, self.display_month, 1),
        )

    def _rebuild_visible_months(self):
        self._day_buttons = {}
        self._day_payload = self._get_payload_map()
        next_year, next_month = _shift_month(self.display_year, self.display_month, 1)
        self.month_label.configure(
            text=(
                f"{GERMAN_MONTH_NAMES[self.display_month]} {self.display_year}  |  "
                f"{GERMAN_MONTH_NAMES[next_month]} {next_year}"
            )
        )

        for month_offset, (month_title, days_frame) in enumerate(self.month_cards):
            year, month = _shift_month(self.display_year, self.display_month, month_offset)
            month_title.configure(text=f"{GERMAN_MONTH_NAMES[month]} {year}")

            for child in days_frame.winfo_children():
                child.destroy()

            month_matrix = pycalendar.Calendar(firstweekday=0).monthdayscalendar(year, month)
            while len(month_matrix) < 6:
                month_matrix.append([0] * 7)

            for row_idx, week in enumerate(month_matrix):
                for col_idx, day in enumerate(week):
                    if not day:
                        ctk.CTkLabel(days_frame, text="", height=40).grid(
                            row=row_idx, column=col_idx, padx=2, pady=2, sticky="nsew"
                        )
                        continue

                    date_key = _date_for_calendar(year, month, day)
                    fg, hover, text = self._pick_colors(date_key)
                    is_selected = date_key == self.selected_date
                    border_color = Theme.ACCENT if is_selected else Theme.BORDER

                    btn = ctk.CTkButton(
                        days_frame,
                        text=str(day),
                        height=40,
                        corner_radius=12,
                        font=_font(12, "bold" if is_selected else "normal"),
                        fg_color=fg,
                        hover_color=hover,
                        text_color=text,
                        border_width=1,
                        border_color=border_color,
                        command=lambda value=date_key: self.select_date(value),
                    )
                    btn.grid(row=row_idx, column=col_idx, padx=2, pady=2, sticky="nsew")
                    self._day_buttons[date_key] = btn

        self._visible_month_keys = self._visible_month_tuple()
        self._payload_revision = self.app.get_calendar_payload_revision()

    def _update_visible_button_styles(self):
        for date_key, btn in self._day_buttons.items():
            fg, hover, text = self._pick_colors(date_key)
            is_selected = date_key == self.selected_date
            btn.configure(
                fg_color=fg,
                hover_color=hover,
                text_color=text,
                border_color=Theme.ACCENT if is_selected else Theme.BORDER,
                font=_font(12, "bold" if is_selected else "normal"),
            )
        self._payload_revision = self.app.get_calendar_payload_revision()

    def _ensure_date_is_visible(self, date_key: str):
        parsed = parse_date(date_key)
        if parsed is None:
            return
        visible_months = {(year, month) for year, month in self._visible_month_tuple()}
        if (parsed.year, parsed.month) not in visible_months:
            self.display_year = parsed.year
            self.display_month = parsed.month

    def refresh(self, force_rebuild: bool = False):
        self._day_payload = self._get_payload_map()
        visible_month_keys = self._visible_month_tuple()
        payload_revision = self.app.get_calendar_payload_revision()
        if force_rebuild or self._visible_month_keys != visible_month_keys or not self._day_buttons:
            self._rebuild_visible_months()
            return
        if payload_revision != self._payload_revision:
            self._update_visible_button_styles()

    def select_date(self, date_key: str):
        now = time.monotonic()
        is_double_click = self._last_click_date == date_key and (now - self._last_click_ts) <= 0.45
        self._last_click_date = date_key
        self._last_click_ts = now

        self.selected_date = date_key
        previous_visible_months = self._visible_month_tuple()
        self._ensure_date_is_visible(date_key)
        self.refresh(force_rebuild=self._visible_month_tuple() != previous_visible_months)
        if self._visible_month_tuple() == previous_visible_months:
            self._update_visible_button_styles()
        if callable(self.on_date_selected):
            self.on_date_selected(date_key, self._day_payload.get(date_key))
        if is_double_click:
            self.activate_date(date_key)

    def activate_date(self, date_key: str):
        if callable(self.on_date_activated):
            self.on_date_activated(date_key, self._day_payload.get(date_key))


# ============================================================
# Pages
# ============================================================
class CalendarPage(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color=Theme.BG)
        self.app = app
        self._selected_date = None
        self._upcoming_cards_window = None
        self._upcoming_last_card_width = 0
        self._upcoming_min_height = 420

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        shell.grid(row=0, column=0, padx=28, pady=28, sticky="nsew")
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(1, weight=1)
        shell.grid_rowconfigure(2, weight=0)

        hero = ctk.CTkFrame(shell, fg_color="transparent")
        hero.grid(row=0, column=0, padx=30, pady=(28, 18), sticky="ew")
        hero.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            hero,
            text="Kalender",
            font=_font(30, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            hero,
            text="Übersicht aller geplanten Touren. Ein Doppelklick öffnet den gewählten Tag in den Liefertouren.",
            font=_font(14),
            text_color=Theme.SUBTEXT,
            justify="left",
        ).grid(row=1, column=0, pady=(8, 0), sticky="w")

        content = ctk.CTkFrame(shell, fg_color="transparent")
        content.grid(row=1, column=0, padx=24, pady=(0, 24), sticky="nsew")
        content.grid_columnconfigure(0, weight=3)
        content.grid_columnconfigure(1, weight=2)
        content.grid_rowconfigure(0, weight=1)

        self.calendar = TourCalendarWidget(
            content,
            app,
            on_date_selected=self._on_date_selected,
            on_date_activated=self._on_date_activated,
        )
        self.calendar.grid(row=0, column=0, sticky="nsew")

        details = ctk.CTkFrame(
            content,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        details.grid(row=0, column=1, padx=(18, 0), sticky="nsew")
        details.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            details,
            text="Legende",
            font=_font(16, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, padx=18, pady=(18, 10), sticky="w")

        self._build_legend_row(details, 1, Theme.WARNING, "1 geplanter Mitarbeitereinsatz")
        self._build_legend_row(details, 2, Theme.DANGER, "2 oder mehr Mitarbeitereinsätze")
        self._build_legend_row(details, 3, Theme.PANEL, "Kein Eintrag für den Tag")

        ctk.CTkLabel(
            details,
            text="Tag",
            font=_font(16, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=4, column=0, padx=18, pady=(22, 6), sticky="w")

        self.selection_label = ctk.CTkLabel(
            details,
            text="Noch kein Datum ausgewählt.",
            font=_font(14),
            text_color=Theme.SUBTEXT,
            justify="left",
            wraplength=320,
        )
        self.selection_label.grid(row=5, column=0, padx=18, pady=(0, 8), sticky="w")

        self.summary_label = ctk.CTkLabel(
            details,
            text="",
            font=_font(13),
            text_color=Theme.SUBTEXT,
            justify="left",
            wraplength=320,
        )
        self.summary_label.grid(row=6, column=0, padx=18, pady=(0, 8), sticky="w")

        self.titles_label = ctk.CTkLabel(
            details,
            text="",
            font=_font(13),
            text_color=Theme.SUBTEXT,
            justify="left",
            wraplength=320,
        )
        self.titles_label.grid(row=7, column=0, padx=18, pady=(0, 18), sticky="w")

        upcoming_shell = ctk.CTkFrame(
            shell,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        upcoming_shell.grid(row=2, column=0, padx=24, pady=(0, 24), sticky="ew")
        upcoming_shell.grid_columnconfigure(0, weight=1)
        upcoming_shell.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            upcoming_shell,
            text="Geplante Touren der nächsten 10 Tage",
            font=_font(16, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, padx=16, pady=(14, 4), sticky="w")

        ctk.CTkLabel(
            upcoming_shell,
            text="5 Tage gleichzeitig sichtbar. Mit Scrollbalken oder Mausrad horizontal weiterblättern.",
            font=_font(12),
            text_color=Theme.SUBTEXT,
        ).grid(row=1, column=0, padx=16, pady=(0, 10), sticky="w")

        self.upcoming_canvas = tk.Canvas(
            upcoming_shell,
            bg=Theme.resolve(Theme.PANEL_2),
            highlightthickness=0,
            bd=0,
            height=self._upcoming_min_height,
            xscrollincrement=16,
        )
        self.upcoming_canvas.grid(row=2, column=0, padx=(14, 14), pady=(0, 8), sticky="nsew")

        self.upcoming_scrollbar = ctk.CTkScrollbar(
            upcoming_shell,
            orientation="horizontal",
            command=self.upcoming_canvas.xview,
            **_scrollbar_kwargs(),
        )
        self.upcoming_scrollbar.grid(row=3, column=0, padx=14, pady=(0, 12), sticky="ew")
        self.upcoming_canvas.configure(xscrollcommand=self.upcoming_scrollbar.set)

        self.upcoming_cards_frame = ctk.CTkFrame(self.upcoming_canvas, fg_color="transparent")
        self._upcoming_cards_window = self.upcoming_canvas.create_window(
            (0, 0),
            window=self.upcoming_cards_frame,
            anchor="nw",
        )
        self.upcoming_cards_frame.bind("<Configure>", self._on_upcoming_frame_configure)
        self.upcoming_canvas.bind("<Configure>", self._on_upcoming_canvas_configure)

        self._bind_horizontal_scroll_mousewheel(self.upcoming_canvas)
        self._bind_horizontal_scroll_mousewheel(self.upcoming_cards_frame)
        self._refresh_upcoming_tours()

    def _build_legend_row(self, master, row: int, color, text: str):
        row_frame = ctk.CTkFrame(master, fg_color="transparent")
        row_frame.grid(row=row, column=0, padx=18, pady=4, sticky="ew")
        row_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkFrame(
            row_frame,
            width=18,
            height=18,
            corner_radius=6,
            fg_color=color,
            border_width=1,
            border_color=Theme.BORDER,
        ).grid(row=0, column=0, padx=(0, 10), sticky="w")

        ctk.CTkLabel(
            row_frame,
            text=text,
            font=_font(13),
            text_color=Theme.SUBTEXT,
            justify="left",
        ).grid(row=0, column=1, sticky="w")

    def _update_selection_details(self, date_key: str | None, payload: dict | None):
        self._selected_date = date_key
        if not date_key:
            self.selection_label.configure(text="Noch kein Datum ausgewählt.")
            self.summary_label.configure(text="")
            self.titles_label.configure(text="")
            return

        self.selection_label.configure(text=f"Ausgewählt: {_display_date_string(date_key)}")

        payload = payload or {}
        tours = int(payload.get("tours", 0) or 0)
        assignments = int(payload.get("assignments", 0) or 0)
        if tours <= 0:
            self.summary_label.configure(text="Für dieses Datum sind keine Touren geplant.")
            self.titles_label.configure(text="")
            return

        self.summary_label.configure(
            text=f"Geplante Touren: {tours} | Mitarbeitereinsätze: {assignments}"
        )
        titles = [str(title).strip() for title in payload.get("titles", []) if str(title).strip()]
        if titles:
            self.titles_label.configure(text="Touren: " + " | ".join(titles[:4]))
        else:
            self.titles_label.configure(text="")

    def _on_date_selected(self, date_key: str, payload: dict | None):
        self._update_selection_details(date_key, payload)

    def _on_date_activated(self, date_key: str, payload: dict | None):
        self._update_selection_details(date_key, payload)
        self.app.set_active_date(date_key)
        self.app.show_page("tours")

    def _bind_horizontal_scroll_mousewheel(self, widget):
        widget.bind("<MouseWheel>", self._on_upcoming_mousewheel, add="+")
        widget.bind("<Shift-MouseWheel>", self._on_upcoming_mousewheel, add="+")
        widget.bind("<Button-4>", self._on_upcoming_mousewheel, add="+")
        widget.bind("<Button-5>", self._on_upcoming_mousewheel, add="+")

    def _on_upcoming_mousewheel(self, event):
        direction = 0
        if getattr(event, "delta", 0):
            direction = -1 if event.delta > 0 else 1
        elif getattr(event, "num", None) == 4:
            direction = -1
        elif getattr(event, "num", None) == 5:
            direction = 1
        if direction == 0:
            return None
        self.upcoming_canvas.xview_scroll(direction * 4, "units")
        return "break"

    def _on_upcoming_frame_configure(self, _event=None):
        self.upcoming_canvas.configure(scrollregion=self.upcoming_canvas.bbox("all"))

    def _on_upcoming_canvas_configure(self, event):
        viewport_width = max(int(getattr(event, "width", 0) or 0), 0)
        if viewport_width <= 0:
            return
        gap = 10
        card_width = max(220, (viewport_width - (gap * 4)) // 5)
        if card_width != self._upcoming_last_card_width:
            self._upcoming_last_card_width = card_width
            self._refresh_upcoming_tours()

    def _resolve_stop_calendar_display(self, stop: dict) -> dict:
        stop_name = str((stop or {}).get("name") or "").strip()
        address = str((stop or {}).get("address") or "").strip()
        product = ""
        for key in ("product", "Produkt", "lieferart", "Lieferart"):
            value = str((stop or {}).get(key) or "").strip()
            if value:
                product = value
                break

        weight_value = self.app._parse_weight((stop or {}).get("weight"))
        marker = self.app._find_marker_for_stop(stop or {})
        if marker is not None:
            data = getattr(marker, "data", {}) or {}
            marker_name = str(data.get("Name", "")).strip()
            street = str(data.get("Strasse", "")).strip()
            plz = str(data.get("PLZ", "")).strip()
            city = str(data.get("Ort", "")).strip()
            marker_address = ", ".join([part for part in [street, " ".join([plz, city]).strip()] if part]).strip()
            if not stop_name:
                stop_name = marker_name
            if not address:
                address = marker_address
            if not product:
                product = _normalize_delivery_type(data.get("Lieferart"))
            if weight_value <= 0:
                weight_value = self.app._parse_weight(data.get("Gewicht"))

        if not stop_name and " / " in address:
            maybe_name, maybe_address = [part.strip() for part in address.split(" / ", 1)]
            if maybe_name and maybe_address:
                stop_name = maybe_name
                address = maybe_address

        if not stop_name:
            stop_name = "Unbekannt"
        if not product:
            product = "Nicht angegeben"
        if not address:
            address = "Adresse unbekannt"

        return {
            "name": stop_name,
            "address": address,
            "product": product,
            "weight": self.app._format_weight(weight_value),
        }

    def _alphabetic_stop_prefix(self, index: int) -> str:
        idx = max(0, int(index))
        chars = []
        while True:
            idx, rem = divmod(idx, 26)
            chars.append(chr(ord("A") + rem))
            if idx == 0:
                break
            idx -= 1
        return "".join(reversed(chars)) + ")"

    def _refresh_upcoming_tours(self):
        for child in self.upcoming_cards_frame.winfo_children():
            child.destroy()

        all_tours = self.app._load_tours()
        today = datetime.now().date()
        end_date = today + timedelta(days=9)

        tours_by_date = {}
        for tour in all_tours:
            date_key = _normalize_date_string(tour.get("date"))
            parsed = parse_date(date_key)
            if not parsed:
                continue
            if parsed < today or parsed > end_date:
                continue
            tours_by_date.setdefault(date_key, []).append(tour)

        canvas_width = max(int(self.upcoming_canvas.winfo_width() or 0), 0)
        gap = 10
        card_width = self._upcoming_last_card_width or max(220, (canvas_width - (gap * 4)) // 5 if canvas_width else 240)

        self.upcoming_cards_frame.grid_rowconfigure(0, weight=1)
        for i in range(10):
            self.upcoming_cards_frame.grid_columnconfigure(i, weight=0, minsize=card_width)
            current_day = today + timedelta(days=i)
            date_key = format_date(current_day)
            day_tours = tours_by_date.get(date_key, [])

            card = ctk.CTkFrame(
                self.upcoming_cards_frame,
                corner_radius=12,
                fg_color=Theme.PANEL,
                border_width=1,
                border_color=Theme.BORDER,
                width=card_width,
            )
            card.grid(row=0, column=i, padx=(0, gap if i < 9 else 0), pady=4, sticky="nsew")
            card.grid_columnconfigure(0, weight=1)

            date_lbl = ctk.CTkLabel(
                card,
                text=current_day.strftime("%d.%m.%Y"),
                font=_font(13, "bold"),
                text_color=Theme.TEXT,
                justify="left",
            )
            date_lbl.grid(row=0, column=0, padx=10, pady=(10, 2), sticky="w")
            self._bind_horizontal_scroll_mousewheel(date_lbl)

            count_lbl = ctk.CTkLabel(
                card,
                text=f"{len(day_tours)} geplante Tour(en)",
                font=_font(12),
                text_color=Theme.SUBTEXT,
                justify="left",
            )
            count_lbl.grid(row=1, column=0, padx=10, pady=(0, 8), sticky="w")
            self._bind_horizontal_scroll_mousewheel(count_lbl)

            self._bind_horizontal_scroll_mousewheel(card)
            content_row = 2
            if not day_tours:
                lbl = ctk.CTkLabel(
                    card,
                    text="Keine Tour geplant.",
                    font=_font(12),
                    text_color=Theme.SUBTEXT,
                    wraplength=max(160, card_width - 22),
                    justify="left",
                )
                lbl.grid(row=content_row, column=0, padx=10, pady=(2, 14), sticky="w")
                self._bind_horizontal_scroll_mousewheel(lbl)
                continue

            for tour in day_tours:
                stops = [stop for stop in (tour.get("stops", []) or []) if isinstance(stop, dict)]
                tour_name = str(tour.get("name") or "").strip() or f"Tour {tour.get('id', '')}".strip()
                tour_header = ctk.CTkLabel(
                    card,
                    text=tour_name,
                    font=_font(12, "bold"),
                    text_color=Theme.ACCENT,
                    wraplength=max(160, card_width - 22),
                    justify="left",
                )
                tour_header.grid(row=content_row, column=0, padx=10, pady=(0, 4), sticky="w")
                self._bind_horizontal_scroll_mousewheel(tour_header)
                content_row += 1

                if not stops:
                    empty_lbl = ctk.CTkLabel(
                        card,
                        text="Keine Stopps gespeichert.",
                        font=_font(11),
                        text_color=Theme.SUBTEXT,
                        justify="left",
                        wraplength=max(160, card_width - 22),
                    )
                    empty_lbl.grid(row=content_row, column=0, padx=10, pady=(0, 14), sticky="w")
                    self._bind_horizontal_scroll_mousewheel(empty_lbl)
                    content_row += 1
                    continue

                for stop_index, stop in enumerate(stops):
                    stop_info = self._resolve_stop_calendar_display(stop)
                    prefix = self._alphabetic_stop_prefix(stop_index)
                    stop_lbl = ctk.CTkLabel(
                        card,
                        text=(
                            f"{prefix} {stop_info['name']}\n"
                            f"   {stop_info['address']}\n"
                            f"   Produkt: {stop_info['product']}\n"
                            f"   Gewicht: {stop_info['weight']}"
                        ),
                        font=_font(11),
                        text_color=Theme.TEXT,
                        justify="left",
                        wraplength=max(160, card_width - 22),
                    )
                    is_last_stop = stop_index == len(stops) - 1
                    bottom_pad = 14 if is_last_stop else 6
                    stop_lbl.grid(row=content_row, column=0, padx=10, pady=(0, bottom_pad), sticky="w")
                    self._bind_horizontal_scroll_mousewheel(stop_lbl)
                    content_row += 1

        self.upcoming_cards_frame.update_idletasks()
        total_width = (card_width * 10) + (gap * 9)
        self.upcoming_canvas.itemconfigure(self._upcoming_cards_window, width=total_width)
        needed_height = max(self._upcoming_min_height, int(self.upcoming_cards_frame.winfo_reqheight()) + 12)
        if int(self.upcoming_canvas.cget("height")) != needed_height:
            self.upcoming_canvas.configure(height=needed_height)
        self.upcoming_canvas.configure(scrollregion=self.upcoming_canvas.bbox("all"))

    def refresh_calendar(self):
        self.calendar.refresh()
        if self._selected_date:
            payload = self.app.get_calendar_payload_map().get(self._selected_date)
            self._update_selection_details(self._selected_date, payload)
        self._refresh_upcoming_tours()

    def refresh(self):
        self.refresh_calendar()


class GPSPage(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color=Theme.BG)
        self.app = app
        self.status_text = tk.StringVar(value="Bereit für native WebView2.")
        self.runtime_hint_text = tk.StringVar(value="")
        self.webview_host = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        topbar = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        topbar.grid(row=0, column=0, padx=20, pady=(20, 12), sticky="ew")
        topbar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            topbar,
            text="GPS",
            font=_font(18, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, padx=16, pady=14, sticky="w")

        ctk.CTkButton(
            topbar,
            text="Neu laden",
            height=36,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self._reload_embedded_view,
        ).grid(row=0, column=1, padx=(8, 0), pady=10)

        ctk.CTkButton(
            topbar,
            text="Im Browser öffnen",
            height=36,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.ACCENT,
            hover_color=Theme.ACCENT_HOVER,
            command=self._open_in_browser,
        ).grid(row=0, column=2, padx=(8, 14), pady=10)

        shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        shell.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(1, weight=1)

        info = ctk.CTkFrame(shell, fg_color="transparent")
        info.grid(row=0, column=0, padx=20, pady=(20, 12), sticky="ew")
        info.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            info,
            text="GPS-Tracking",
            font=_font(20, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            info,
            text=(
                "Das GPS-Portal wird direkt in dieser Seite in einer eingebetteten nativen WebView2-Ansicht geladen. "
                "Login, Cookies und moderne Browser-Funktionen bleiben dabei in einer echten Edge/WebView2-Laufzeit."
            ),
            font=_font(13),
            text_color=Theme.SUBTEXT,
            justify="left",
            wraplength=980,
        ).grid(row=1, column=0, pady=(8, 12), sticky="w")

        url_row = ctk.CTkFrame(info, fg_color="transparent")
        url_row.grid(row=2, column=0, sticky="ew")
        url_row.grid_columnconfigure(0, weight=1)

        self.url_entry = ctk.CTkEntry(
            url_row,
            height=38,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            border_color=Theme.BORDER,
            text_color=Theme.TEXT,
        )
        self.url_entry.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        self.url_entry.insert(0, GPS_TRACKING_URL)
        self.url_entry.configure(state="readonly")

        ctk.CTkButton(
            url_row,
            text="Kopieren",
            width=110,
            height=38,
            corner_radius=12,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self._copy_url,
        ).grid(row=0, column=1, sticky="e")

        body = ctk.CTkFrame(
            shell,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        body.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(0, weight=1)

        browser_panel = ctk.CTkFrame(
            body,
            corner_radius=16,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        browser_panel.grid(row=0, column=0, padx=16, pady=16, sticky="nsew")
        browser_panel.grid_columnconfigure(0, weight=1)
        browser_panel.grid_rowconfigure(0, weight=1)

        embedded_host_shell = ctk.CTkFrame(
            browser_panel,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        embedded_host_shell.grid(row=0, column=0, padx=16, pady=16, sticky="nsew")
        embedded_host_shell.grid_columnconfigure(0, weight=1)
        embedded_host_shell.grid_rowconfigure(0, weight=1)

        self.webview_host = EmbeddedWebView2Frame(
            embedded_host_shell,
            self.app,
            url=GPS_TRACKING_URL,
            status_callback=self._update_status,
        )
        self.webview_host.grid(row=0, column=0, sticky="nsew")

    def _update_status(self, message: str):
        self.status_text.set(str(message or "").strip() or "Bereit")

    def _refresh_runtime_status(self):
        runtime_path = _resolve_webview2_runtime_path(self.app.base_dir)
        if runtime_path:
            if _is_runtime_inside_bundle(self.app.base_dir, runtime_path):
                self.runtime_hint_text.set(f"Gebundene Fixed-Version-Runtime: {runtime_path}")
            else:
                self.runtime_hint_text.set(f"Systeminstallierte WebView2-Runtime: {runtime_path}")
        else:
            self.runtime_hint_text.set(
                "Keine gebundene Fixed-Version-Runtime gefunden. Fallback auf System-WebView2."
            )
        if self.webview_host is None:
            self._update_status("Bereit für eingebettete WebView2.")

    def _reload_embedded_view(self):
        if self.webview_host is None:
            return
        self._update_status("GPS-Seite wird neu geladen...")
        self.webview_host.reload()

    def _open_in_browser(self):
        try:
            webbrowser.open(GPS_TRACKING_URL)
        except Exception as exc:
            messagebox.showerror("GPS", f"Die GPS-Seite konnte nicht geöffnet werden:\n{exc}")

    def _copy_url(self):
        self.clipboard_clear()
        self.clipboard_append(GPS_TRACKING_URL)
        messagebox.showinfo("GPS", "Die GPS-URL wurde in die Zwischenablage kopiert.")

    def refresh(self):
        try:
            self.url_entry.configure(state="normal")
            self.url_entry.delete(0, "end")
            self.url_entry.insert(0, GPS_TRACKING_URL)
            self.url_entry.configure(state="readonly")
        except Exception:
            pass
        self._refresh_runtime_status()
        if self.webview_host is not None:
            self.webview_host.navigate(GPS_TRACKING_URL)

    def on_show(self):
        self.refresh()

    def on_hide(self):
        return


class StartMenuPage(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color=Theme.BG)
        self.app = app
        self._update_hint_label = None
        self._logo_image = self._load_logo_image()

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        shell.grid(row=0, column=0, padx=28, pady=28, sticky="nsew")
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(1, weight=1)

        hero = ctk.CTkFrame(shell, fg_color="transparent")
        hero.grid(row=0, column=0, padx=30, pady=(30, 16), sticky="ew")
        hero.grid_columnconfigure(0, weight=1)

        if self._logo_image is not None:
            ctk.CTkLabel(hero, text="", image=self._logo_image).grid(row=0, column=0, pady=(0, 18))

        title = ctk.CTkLabel(
            hero, text="GAWELA Tourenplaner", font=_font(32, "bold"), text_color=Theme.TEXT
        )
        title.grid(row=1, column=0, sticky="n")

        subtitle = ctk.CTkLabel(
            hero,
            text="Startseite mit Direktzugriff auf alle Bereiche.",
            font=_font(15),
            text_color=Theme.SUBTEXT,
            justify="center",
        )
        subtitle.grid(row=2, column=0, pady=(8, 0), sticky="n")

        launcher = ctk.CTkFrame(
            shell,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        launcher.grid(row=1, column=0, padx=24, pady=(0, 18), sticky="nsew")
        launcher.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(
            launcher,
            text="Bereiche",
            font=_font(18, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, columnspan=2, padx=20, pady=(20, 6), sticky="w")

        ctk.CTkLabel(
            launcher,
            text="Die linke Navigation wird erst nach dem Wechsel in einen Bereich eingeblendet.",
            font=_font(13),
            text_color=Theme.SUBTEXT,
            justify="left",
        ).grid(row=1, column=0, columnspan=2, padx=20, pady=(0, 12), sticky="w")

        for index, (page_name, label) in enumerate(self._navigation_items()):
            row = (index // 2) + 2
            column = index % 2
            button = ctk.CTkButton(
                launcher,
                text=label,
                height=58,
                corner_radius=16,
                font=_font(15, "bold"),
                fg_color=Theme.PANEL,
                hover_color=Theme.BORDER,
                text_color=Theme.TEXT,
                command=lambda name=page_name: app.show_page(name),
            )
            padx = (20, 10) if column == 0 else (10, 20)
            button.grid(row=row, column=column, padx=padx, pady=10, sticky="ew")

        utility = ctk.CTkFrame(shell, fg_color="transparent")
        utility.grid(row=2, column=0, padx=24, pady=(0, 24), sticky="ew")
        utility.grid_columnconfigure(0, weight=1)
        utility.grid_columnconfigure(1, weight=0)

        info_card = ctk.CTkFrame(
            utility,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        info_card.grid(row=0, column=0, sticky="ew")
        info_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            info_card,
            text="Hinweis",
            font=_font(15, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, padx=16, pady=(14, 4), sticky="w")

        self._update_hint_label = ctk.CTkLabel(
            info_card,
            text="Update-Status wird geladen ...",
            font=_font(13),
            text_color=Theme.SUBTEXT,
            justify="left",
            wraplength=780,
        )
        self._update_hint_label.grid(row=1, column=0, padx=16, pady=(0, 8), sticky="w")

        ctk.CTkLabel(
            info_card,
            text="Von hier aus lassen sich Karte, Aufträge, Touren, Mitarbeiter, Fahrzeuge und Einstellungen direkt öffnen.",
            font=_font(13),
            text_color=Theme.SUBTEXT,
            justify="left",
            wraplength=780,
        ).grid(row=2, column=0, padx=16, pady=(0, 14), sticky="w")

        self.btn_darkmode = ctk.CTkButton(
            utility,
            text=self._dark_btn_text(),
            height=40,
            corner_radius=14,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self._toggle_dark,
        )
        self.btn_darkmode.grid(row=0, column=1, padx=(16, 0), sticky="e")

    def _dark_btn_text(self):
        mode = ctk.get_appearance_mode()
        return "Darkmode: AUS" if mode == "Dark" else "Darkmode: EIN"

    def _load_logo_image(self):
        path = getattr(self.app, "app_logo_path", "")
        if not path or not os.path.exists(path):
            return None
        try:
            image = Image.open(path)
            return ctk.CTkImage(light_image=image, dark_image=image, size=(240, 240))
        except Exception:
            logger.exception("Start page logo could not be loaded.")
            return None

    def _navigation_items(self):
        items = [
            ("calendar", "Kalender"),
            ("map", "Karte"),
            ("gps", "GPS"),
            ("list", "Auftragsliste"),
            ("nonmap", "Nicht-Karten-Aufträge"),
            ("tours", "Liefertouren"),
            ("employees", "Mitarbeiter"),
            ("vehicles", "Fahrzeuge"),
            ("settings", "Einstellungen"),
        ]
        if SHOW_UPDATE_PAGE_IN_MENU:
            items.append(("update", "Updates"))
        return items

    def _toggle_dark(self):
        self.app.toggle_darkmode()
        self.btn_darkmode.configure(text=self._dark_btn_text())

    def on_theme_changed(self):
        try:
            self.btn_darkmode.configure(text=self._dark_btn_text())
        except Exception:
            pass

    def refresh(self):
        if self._update_hint_label is not None and hasattr(self.app, "get_startup_update_message"):
            self._update_hint_label.configure(text=self.app.get_startup_update_message())


class MapPage(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color=Theme.BG)
        self.app = app

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Topbar
        topbar = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        topbar.grid(row=0, column=0, padx=20, pady=(20, 12), sticky="ew")
        topbar.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(
            topbar,
            text="Karte",
            font=_font(18, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, padx=16, pady=14, sticky="w")

        # Search
        search_wrap = ctk.CTkFrame(topbar, corner_radius=14, fg_color=Theme.PANEL_2)
        search_wrap.grid(row=0, column=2, padx=14, pady=10, sticky="ew")
        search_wrap.grid_columnconfigure(0, weight=1)

        search_entry = ctk.CTkEntry(
            search_wrap,
            placeholder_text="Ort suchen (Enter)…",
            height=36,
            corner_radius=12,
        )
        search_entry.grid(row=0, column=0, padx=(10, 8), pady=10, sticky="ew")
        search_entry.bind("<Return>", lambda e: app.search_location())

        ctk.CTkButton(
            search_wrap,
            text="Suchen",
            height=36,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.ACCENT,
            hover_color=Theme.ACCENT_HOVER,
            command=app.search_location,
        ).grid(row=0, column=1, padx=(0, 10), pady=10)

        status_filter = ctk.CTkButton(
            topbar,
            text=app._map_filter_button_text(),
            height=36,
            width=210,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=app.open_map_filter_dialog,
            anchor="w",
        )
        status_filter.grid(row=0, column=3, padx=(0, 14), pady=10, sticky="e")

        ctk.CTkButton(
            topbar,
            text="Route exportieren",
            height=36,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.SUCCESS,
            hover_color=Theme.SUCCESS_HOVER,
            command=app.export_route,
        ).grid(row=0, column=4, padx=(0, 14), pady=10, sticky="e")

        # Main area
        body = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        body.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=0)
        body.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body, fg_color="transparent")
        left.grid(row=0, column=0, padx=14, pady=14, sticky="nsew")
        left.grid_rowconfigure(1, weight=1)
        left.grid_columnconfigure(0, weight=1)

        options = ctk.CTkFrame(left, fg_color="transparent")
        options.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        options.grid_columnconfigure(0, weight=1)

        legend_section = DropdownSection(options, "Legende", default_open=True)
        legend_section.grid(row=0, column=0, sticky="ew", pady=(0, 0))
        legend_section.body.grid_columnconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8), weight=1)

        ctk.CTkLabel(
            legend_section.body,
            text="Farben (Auftragsstatus)",
            text_color=Theme.SUBTEXT,
            font=_font(12, "bold"),
        ).grid(row=0, column=0, columnspan=6, padx=8, pady=(6, 2), sticky="w")

        ctk.CTkLabel(
            legend_section.body,
            text="Formen (Lieferart)",
            text_color=Theme.SUBTEXT,
            font=_font(12, "bold"),
        ).grid(row=0, column=6, columnspan=3, padx=8, pady=(6, 2), sticky="w")

        legend_icons = []
        status_legend_frame = ctk.CTkFrame(legend_section.body, fg_color="transparent")
        status_legend_frame.grid(row=1, column=0, columnspan=6, padx=8, pady=(0, 6), sticky="ew")
        status_legend_frame.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)

        status_items = [
            ("nicht festgelegt", app.status_colors.get("nicht festgelegt", "#575757")),
            ("Bestellt", app.status_colors.get("Bestellt", "#5959FF")),
            ("Auf dem Weg", app.status_colors.get("Auf dem Weg", "#D17E0D")),
            ("Im Lager", app.status_colors.get("Im Lager", "#5CFF59")),
            ("Bereits eingeplant", app.status_colors.get("Bereits eingeplant", "#94A3B8")),
            ("Tourfarbe", app.tour_pin_color),
        ]
        status_columns = 6
        for item_idx, (label_text, color) in enumerate(status_items):
            icon = app._make_circle_icon(color, 16)
            legend_icons.append(icon)
            row = ctk.CTkFrame(status_legend_frame, fg_color="transparent")
            row_idx = item_idx // status_columns
            col_idx = item_idx % status_columns
            row.grid(row=row_idx, column=col_idx, sticky="w", pady=1, padx=(0, 12))
            ctk.CTkLabel(row, text="", image=icon).pack(side="left", padx=(0, 8))
            ctk.CTkLabel(row, text=label_text, font=_font(12), text_color=Theme.TEXT).pack(side="left")

        delivery_legend_frame = ctk.CTkFrame(legend_section.body, fg_color="transparent")
        delivery_legend_frame.grid(row=1, column=6, columnspan=3, padx=8, pady=(0, 6), sticky="ew")
        delivery_legend_frame.grid_columnconfigure((0, 1, 2), weight=1)

        delivery_columns = 3
        for item_idx, delivery_type in enumerate(DELIVERY_TYPE_OPTIONS):
            normalized = _normalize_delivery_type(delivery_type)
            icon = app._make_delivery_icon(normalized, app.status_colors.get("Bestellt", "#5959FF"), 16)
            legend_icons.append(icon)
            row = ctk.CTkFrame(delivery_legend_frame, fg_color="transparent")
            row_idx = item_idx // delivery_columns
            col_idx = item_idx % delivery_columns
            row.grid(row=row_idx, column=col_idx, sticky="w", pady=1, padx=(0, 12))
            ctk.CTkLabel(row, text="", image=icon).pack(side="left", padx=(0, 8))
            ctk.CTkLabel(row, text=normalized, font=_font(12), text_color=Theme.TEXT).pack(side="left")

        app.map_legend_icons = legend_icons

        # -------- ROUTE PANEL + MAP (resizable splitter) --------
        split_bg = Theme.resolve(Theme.BORDER)
        map_split = tk.PanedWindow(
            left,
            orient="horizontal",
            sashwidth=8,
            showhandle=False,
            bd=0,
            relief="flat",
            background=split_bg,
            sashrelief="flat",
        )
        map_split.grid(row=1, column=0, sticky="nsew")

        route_panel = ctk.CTkFrame(map_split, corner_radius=16, fg_color=Theme.PANEL_2, width=560)
        route_panel.grid_propagate(False)
        route_panel.grid_rowconfigure(2, weight=1)
        route_panel.grid_columnconfigure(0, weight=1)

        header_row = ctk.CTkFrame(route_panel, fg_color="transparent")
        header_row.grid(row=0, column=0, padx=12, pady=(12, 6), sticky="ew")
        header_info = ctk.CTkFrame(header_row, fg_color="transparent")
        header_info.pack(side="left", fill="x", expand=True)

        header_actions = ctk.CTkFrame(header_row, fg_color="transparent")
        header_actions.pack(side="right", anchor="ne")

        ctk.CTkLabel(
            header_info,
            text="Aktuelle Route",
            font=_font(15, "bold"),
            text_color=Theme.TEXT,
        ).pack(anchor="w")

        tour_label = ctk.CTkLabel(
            header_info,
            text="(manuell)",
            font=_font(12),
            text_color=Theme.SUBTEXT,
        )
        tour_label.pack(anchor="w", pady=(2, 0))

        btn_pick_tour = ctk.CTkButton(
            header_actions,
            text="Tour wählen ▾",
            height=32,
            corner_radius=12,
            font=_font(12, "bold"),
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: app.open_tour_picker(btn_pick_tour),
        )
        btn_pick_tour.pack(side="left", padx=(0, 6))

        btn_edit_tour = ctk.CTkButton(
            header_actions,
            text="Bearbeiten",
            height=32,
            corner_radius=12,
            font=_font(12, "bold"),
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=app.open_edit_current_tour,
        )
        btn_edit_tour.pack(side="left", padx=(0, 6))

        btn_new_tour = ctk.CTkButton(
            header_actions,
            text="+",
            width=36,
            height=32,
            corner_radius=12,
            font=_font(14, "bold"),
            fg_color=Theme.SUCCESS,
            hover_color=Theme.SUCCESS_HOVER,
            text_color=("white", "white"),
            command=app.open_create_tour_dialog,
        )
        btn_new_tour.pack(side="left")

        btn_clear_tour = ctk.CTkButton(
            header_actions,
            text="x",
            width=36,
            height=32,
            corner_radius=12,
            font=_font(14, "bold"),
            fg_color=Theme.MUTED_BTN,
            hover_color=Theme.MUTED_BTN_HOVER,
            text_color=("white", "white"),
            command=app.unload_current_tour,
        )
        btn_clear_tour.pack(side="left", padx=(6, 0))

        app.lbl_current_tour = tour_label

        start_time_row = ctk.CTkFrame(route_panel, fg_color="transparent")
        start_time_row.grid(row=1, column=0, padx=10, pady=(0, 6), sticky="ew")
        start_time_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            start_time_row,
            text="Startzeit",
            font=_font(12, "bold"),
            text_color=Theme.SUBTEXT,
        ).grid(row=0, column=0, padx=(2, 8), pady=2, sticky="w")

        start_time_entry = TimeInput(start_time_row, width=110, height=32)
        start_time_entry.grid(row=0, column=1, padx=(0, 8), pady=2, sticky="w")
        start_time_entry.set("08:00")

        ctk.CTkButton(
            start_time_row,
            text="Übernehmen",
            width=110,
            height=32,
            corner_radius=12,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: app.apply_route_start_time(start_time_entry.get()),
        ).grid(row=0, column=2, pady=2, sticky="e")

        ctk.CTkButton(
            start_time_row,
            text="Route optimieren",
            width=150,
            height=32,
            corner_radius=12,
            font=_font(12, "bold"),
            fg_color=Theme.SUCCESS,
            hover_color=Theme.SUCCESS_HOVER,
            text_color=("white", "white"),
            command=app.optimize_current_route_order,
        ).grid(row=0, column=3, padx=(8, 0), pady=2, sticky="e")

        tv_shell = ctk.CTkFrame(route_panel, corner_radius=12, fg_color=Theme.PANEL)
        tv_shell.grid(row=2, column=0, padx=10, pady=(6, 6), sticky="nsew")
        tv_shell.grid_rowconfigure(0, weight=1)
        tv_shell.grid_columnconfigure(0, weight=1)

        route_tree = ttk.Treeview(
            tv_shell,
            columns=("Pos", "Adresse", "Fenster", "Service", "ETA", "ETD", "Gewicht"),
            show="headings",
            selectmode="browse",
        )
        route_tree.heading("Pos", text="#")
        route_tree.heading("Adresse", text="Adresse / Name")
        route_tree.heading("Fenster", text="Fenster")
        route_tree.heading("Service", text="Aufenthalt")
        route_tree.heading("ETA", text="ETA")
        route_tree.heading("ETD", text="ETD")
        route_tree.heading("Gewicht", text="Gewicht")
        route_tree.column("Pos", width=40, anchor="center", stretch=False)
        route_tree.column("Adresse", width=210, anchor="w", stretch=True)
        route_tree.column("Fenster", width=105, anchor="center", stretch=False)
        route_tree.column("Service", width=86, anchor="center", stretch=False)
        route_tree.column("ETA", width=74, anchor="center", stretch=False)
        route_tree.column("ETD", width=74, anchor="center", stretch=False)
        route_tree.column("Gewicht", width=85, anchor="e", stretch=False)
        route_tree.grid(row=0, column=0, sticky="nsew")

        app.route_tree = route_tree
        route_tree.bind("<ButtonPress-1>", app.on_route_tree_drag_start, add="+")
        route_tree.bind("<B1-Motion>", app.on_route_tree_drag_motion, add="+")
        route_tree.bind("<ButtonRelease-1>", app.on_route_tree_drag_release, add="+")

        route_total = ctk.CTkLabel(route_panel, text="Totalgewicht: 0", font=_font(13, "bold"), text_color=Theme.TEXT)
        route_total.grid(row=3, column=0, padx=12, pady=(0, 4), sticky="w")
        app.route_total_label = route_total

        schedule_summary = ctk.CTkLabel(
            route_panel,
            text="Zeitplan: noch nicht berechnet",
            font=_font(12),
            text_color=Theme.SUBTEXT,
            justify="left",
        )
        schedule_summary.grid(row=4, column=0, padx=12, pady=(0, 6), sticky="w")

        segments_shell = ctk.CTkFrame(route_panel, corner_radius=12, fg_color=Theme.PANEL)
        segments_shell.grid(row=5, column=0, padx=10, pady=(0, 8), sticky="ew")
        segments_shell.grid_columnconfigure(0, weight=1)
        segments_shell.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            segments_shell,
            text="Fahrzeiten",
            font=_font(13, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, padx=10, pady=(10, 4), sticky="w")

        segment_list = ctk.CTkScrollableFrame(
            segments_shell,
            height=150,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            **_scrollable_frame_kwargs(),
        )
        segment_list.grid(row=1, column=0, padx=8, pady=(0, 8), sticky="ew")
        segment_list.grid_columnconfigure(0, weight=1)

        btns = ctk.CTkFrame(route_panel, fg_color="transparent")
        btns.grid(row=6, column=0, padx=10, pady=(0, 12), sticky="ew")
        btns.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkButton(
            btns,
            text="↑",
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=app.move_selected_stop_up,
        ).grid(row=0, column=0, padx=(0, 6), sticky="ew")

        ctk.CTkButton(
            btns,
            text="↓",
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=app.move_selected_stop_down,
        ).grid(row=0, column=1, padx=6, sticky="ew")

        ctk.CTkButton(
            btns,
            text="Zeitfenster",
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=app.open_selected_stop_editor,
        ).grid(row=0, column=2, padx=6, sticky="ew")

        ctk.CTkButton(
            btns,
            text="Entfernen",
            height=36,
            corner_radius=12,
            fg_color=Theme.DANGER,
            hover_color=Theme.DANGER_HOVER,
            command=app.remove_selected_stop_from_route_panel,
        ).grid(row=0, column=3, padx=(6, 0), sticky="ew")

        route_tree.bind("<Double-1>", app.on_route_tree_double_click)

        map_container = ctk.CTkFrame(map_split, corner_radius=16, fg_color=Theme.PANEL_2)
        map_container.grid_rowconfigure(0, weight=1)
        map_container.grid_columnconfigure(0, weight=1)

        map_split.add(route_panel, minsize=460, width=560)
        map_split.add(map_container, minsize=480)

        map_widget = tkintermapview.TkinterMapView(map_container, corner_radius=14)
        map_widget.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        map_widget.set_position(46.8182, 8.2275)
        map_widget.set_zoom(8)
        map_widget.set_tile_server(
            "https://mt0.google.com/vt/lyrs=m&hl=en&x={x}&y={y}&z={z}&s=Ga",
            max_zoom=22,
        )

        info_card = ctk.CTkFrame(
            body,
            width=320,
            corner_radius=18,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        info_card.grid(row=0, column=1, padx=(0, 14), pady=14, sticky="nsew")
        info_card.grid_columnconfigure(0, weight=1)
        info_card.grid_columnconfigure(1, weight=0)
        info_card.grid_rowconfigure(1, weight=1)
        info_card.grid_forget()

        info_title = ctk.CTkLabel(info_card, text="Details", font=_font(16, "bold"), text_color=Theme.TEXT)
        info_title.grid(row=0, column=0, padx=14, pady=(14, 6), sticky="w")

        info_label = tk.Text(
            info_card,
            wrap="word",
            height=1,
            font=_font(13),
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            padx=0,
            pady=0,
            bg=Theme.resolve(Theme.PANEL_2),
            fg=Theme.resolve(Theme.TEXT),
            insertbackground=Theme.resolve(Theme.TEXT),
        )
        info_label.tag_configure("bold", font=_font(13, "bold"))
        info_scrollbar = ctk.CTkScrollbar(
            info_card,
            orientation="vertical",
            command=info_label.yview,
            **_scrollbar_kwargs(),
        )
        info_label.configure(yscrollcommand=info_scrollbar.set)
        info_label.configure(state="disabled")
        info_label.grid(row=1, column=0, padx=(14, 6), pady=(0, 10), sticky="nsew")
        info_scrollbar.grid(row=1, column=1, padx=(0, 14), pady=(0, 10), sticky="ns")

        ctk.CTkButton(
            info_card,
            text="Zu Tour hinzufügen",
            height=36,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.SUCCESS,
            hover_color=Theme.SUCCESS_HOVER,
            command=app.add_current_to_route,
        ).grid(row=2, column=0, padx=14, pady=(0, 8), sticky="ew")

        ctk.CTkButton(
            info_card,
            text="E-Mail senden",
            height=36,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.ACCENT,
            hover_color=Theme.ACCENT_HOVER,
            command=app.open_email_client,
        ).grid(row=3, column=0, padx=14, pady=(0, 10), sticky="ew")

        # Button "Auftrag bearbeiten"
        btn_edit_order = ctk.CTkButton(
            info_card,
            text="Auftrag bearbeiten",
            height=36,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: app.open_customer_editor(app.current_selected_marker),
        )
        btn_edit_order.grid(row=4, column=0, padx=14, pady=(0, 10), sticky="ew")
        btn_edit_order.grid_remove()
        app.btn_edit_order = btn_edit_order

        # WICHTIGER FIX:
        # Label und OptionMenu dürfen NICHT in derselben Grid-Zeile liegen,
        # sonst überdeckt das Label das Dropdown (klickt dann "nicht").
        ctk.CTkLabel(
            info_card,
            text="Auftragsstatus",
            text_color=Theme.SUBTEXT,
            font=_font(13),
        ).grid(row=5, column=0, padx=14, pady=(6, 0), sticky="w")

        status_menu = ctk.CTkOptionMenu(
            info_card,
            values=["nicht festgelegt", "Bestellt", "Auf dem Weg", "Im Lager"],
            command=app.set_selected_pin_status,
            corner_radius=12,
            height=36,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL,
            button_color=Theme.ACCENT,
            button_hover_color=Theme.ACCENT_HOVER,
            text_color=Theme.TEXT,
        )
        status_menu.grid(row=6, column=0, padx=14, pady=8, sticky="ew")

        btn_show_tour = ctk.CTkButton(
            info_card,
            text="Tour anzeigen",
            height=36,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=app.show_tour_of_selected_pin,
        )
        btn_remove_from_tour = ctk.CTkButton(
            info_card,
            text="Aus Tour entfernen",
            height=36,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.DANGER,
            hover_color=Theme.DANGER_HOVER,
            command=app.remove_selected_pin_from_tour,
        )
        btn_show_tour.grid(row=7, column=0, padx=14, pady=(6, 6), sticky="ew")
        btn_remove_from_tour.grid(row=8, column=0, padx=14, pady=(0, 10), sticky="ew")
        btn_show_tour.grid_remove()
        btn_remove_from_tour.grid_remove()

        ctk.CTkButton(
            info_card,
            text="Schließen",
            height=32,
            corner_radius=12,
            fg_color=Theme.MUTED_BTN,
            hover_color=Theme.MUTED_BTN_HOVER,
            command=app.hide_info_card,
        ).grid(row=9, column=0, padx=14, pady=(0, 14), sticky="ew")

        # Expose to app
        app.main_frame = body
        app.map_container = map_container
        app.info_card = info_card
        app.info_label = info_label
        app.search_entry = search_entry
        app.map_widget = map_widget
        app.map_split = map_split
        app.status_menu = status_menu
        app.btn_show_tour = btn_show_tour
        app.btn_remove_from_tour = btn_remove_from_tour
        app.lbl_route_employee_summary = None
        app.lbl_route_vehicle_summary = None
        app.lbl_route_trailer_summary = None
        app.btn_edit_current_tour = btn_edit_tour
        app.route_start_time_entry = start_time_entry
        app.route_segment_list = segment_list
        app.route_schedule_summary_label = schedule_summary
        app.map_filter_button = status_filter

    def on_theme_changed(self):
        try:
            self.app.map_split.configure(background=Theme.resolve(Theme.BORDER))
        except Exception:
            pass


class OrdersListPage(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color=Theme.BG)
        self.app = app
        self._filter_job = None

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        topbar = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        topbar.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 12))
        topbar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(topbar, text="Auftragsliste", font=_font(18, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )

        ctk.CTkButton(
            topbar,
            text="Aktualisieren",
            width=120,
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.refresh,
        ).grid(row=0, column=2, padx=12, pady=10)

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="w", padx=20, pady=(0, 6))
        actions.grid_columnconfigure((0, 1), weight=0)

        ctk.CTkButton(
            actions,
            text="SQL-Datenordner wählen",
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.app.select_sql_data_dir,
        ).grid(row=0, column=0, padx=(0, 8), pady=2, sticky="w")

        self.sql_import_button = ctk.CTkButton(
            actions,
            text="SQL importieren",
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.app.import_from_sql,
        )
        self.sql_import_button.grid(row=0, column=1, padx=(8, 0), pady=2, sticky="w")

        self.import_status_label = ctk.CTkLabel(
            actions,
            text="Importstatus: Bereit",
            anchor="w",
            text_color=Theme.SUBTEXT,
            font=_font(12),
        )
        self.import_status_label.grid(row=1, column=0, columnspan=2, padx=(0, 0), pady=(4, 0), sticky="w")

        self.search_var = tk.StringVar(value="")
        search_entry = ctk.CTkEntry(
            topbar,
            textvariable=self.search_var,
            placeholder_text="Suchen (z.B. Auftragsnummer)...",
            width=320,
            height=36,
            corner_radius=12,
        )
        search_entry.grid(row=0, column=1, padx=(10, 8), pady=10, sticky="e")
        search_entry.bind("<KeyRelease>", self._schedule_filter, add="+")

        ctk.CTkButton(
            topbar,
            text="Adresse löschen",
            width=140,
            height=36,
            corner_radius=12,
            fg_color=Theme.DANGER,
            hover_color=Theme.DANGER_HOVER,
            command=self.delete_selected_order,
        ).grid(row=0, column=3, padx=(0, 14), pady=10)

        self.info = ctk.CTkLabel(self, text="", anchor="w", text_color=Theme.SUBTEXT, font=_font(13))
        self.info.grid(row=1, column=0, padx=24, pady=(0, 6), sticky="w")

        table_shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        table_shell.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 20))
        table_shell.grid_columnconfigure(0, weight=1)
        table_shell.grid_rowconfigure(0, weight=1)

        self.columns = [
            "Auftragsnummer",
            "Bestelldatum",
            "Auftragsadresse",
            "Lieferadresse",
            "Email",
            "Telefon",
            "Gewicht",
            "Produkte",
            "Status",
        ]
        self.tree = ttk.Treeview(table_shell, columns=self.columns, show="headings")
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, anchor="w", stretch=True)
        self.tree.column("Auftragsadresse", width=260, anchor="w", stretch=True)
        self.tree.column("Lieferadresse", width=280, anchor="w", stretch=True)
        self.tree.column("Produkte", width=320, anchor="w", stretch=True)

        vsb = ctk.CTkScrollbar(table_shell, orientation="vertical", command=self.tree.yview, **_scrollbar_kwargs())
        hsb = ctk.CTkScrollbar(table_shell, orientation="horizontal", command=self.tree.xview, **_scrollbar_kwargs())
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        vsb.grid(row=0, column=1, sticky="ns", padx=(0, 12), pady=12)
        hsb.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 12))

        self.tree.bind("<Double-1>", self.on_double_click)

        self._all_rows = []
        self._order_to_marker = {}
        self._row_to_entry = {}
        self.update_import_status()

    def _schedule_filter(self, _event=None):
        if self._filter_job is not None:
            try:
                self.after_cancel(self._filter_job)
            except Exception:
                pass
        self._filter_job = self.after(130, self.apply_filter)

    def refresh(self):
        self.update_import_status()
        self._all_rows = []
        self._order_to_marker = {}
        self._row_to_entry = {}
        for m in getattr(self.app, "marker_list", []):
            if getattr(m, "is_system", False):
                continue
            data = getattr(m, "data", {}) or {}
            if data:
                d = _normalize_order_and_delivery_addresses(dict(data))
                d["Bestelldatum"] = _display_date_string(d.get("Bestelldatum")) or str(d.get("Bestelldatum") or "")
                d["Status"] = getattr(m, "status", d.get("Status", "nicht festgelegt"))
                d["Auftragsadresse"] = _format_order_address(d)
                d["Lieferadresse"] = _format_delivery_address(d)
                auftrag = str(d.get("Auftragsnummer", "")).strip()
                if auftrag:
                    self._order_to_marker[auftrag] = m
                search_text = " ".join(str(d.get(col, "")) for col in self.columns).lower()
                identity = self.app._item_identity_key(d)
                self._all_rows.append((d, search_text, {"type": "marker", "marker": m, "identity": identity}))
        for pending in getattr(self.app, "pending_sql_orders", []):
            if not isinstance(pending, dict):
                continue
            d = _normalize_order_and_delivery_addresses(dict(pending))
            d["Bestelldatum"] = _display_date_string(d.get("Bestelldatum")) or str(d.get("Bestelldatum") or "")
            d.setdefault("Status", "nicht festgelegt")
            d.setdefault("Lieferart", DEFAULT_DELIVERY_TYPE)
            d["Auftragsadresse"] = _format_order_address(d)
            d["Lieferadresse"] = _format_delivery_address(d)
            search_text = " ".join(str(d.get(col, "")) for col in self.columns).lower()
            identity = self.app._item_identity_key(d)
            self._all_rows.append((d, search_text, {"type": "pending", "marker": None, "identity": identity}))
        self.apply_filter()

    def clear_search(self):
        self.search_var.set("")
        self.apply_filter()

    def _get_selected_marker(self):
        selection = self.tree.selection()
        item_id = selection[0] if selection else self.tree.focus()
        if not item_id:
            return None, None, None

        entry = self._row_to_entry.get(item_id)
        if isinstance(entry, dict):
            marker = entry.get("marker")
            if marker is not None:
                auftrag = str(getattr(marker, "auftragsnummer", "")).strip()
            else:
                values = self.tree.item(item_id, "values")
                auftrag = str(values[0]).strip() if values else ""
            return marker, auftrag, entry

        values = self.tree.item(item_id, "values")
        if not values:
            return None, None, None

        auftrag = str(values[0]).strip()
        if not auftrag:
            return None, None, None

        marker = self._order_to_marker.get(auftrag)
        if marker is None:
            for m in getattr(self.app, "marker_list", []):
                if getattr(m, "is_system", False):
                    continue
                if str(getattr(m, "auftragsnummer", "")).strip() == auftrag:
                    marker = m
                    self._order_to_marker[auftrag] = m
                    break
        return marker, auftrag, {"type": "marker" if marker else "pending", "marker": marker, "identity": None}

    def delete_selected_order(self):
        marker, auftrag, entry = self._get_selected_marker()
        if marker is None and not isinstance(entry, dict):
            messagebox.showwarning("Auftrag löschen", "Bitte zuerst einen Auftrag in der Liste auswählen.")
            return

        label = auftrag or "ohne Auftragsnummer"
        if not messagebox.askyesno(
            "Auftrag löschen",
            f"Soll der ausgewählte Auftrag wirklich gelöscht werden?\n\nAuftragsnummer: {label}",
        ):
            return

        if marker is None:
            identity = entry.get("identity") if isinstance(entry, dict) else None
            if identity is None:
                messagebox.showwarning("Auftrag löschen", "Ausgewählter Pending-Auftrag konnte nicht aufgelöst werden.")
                return
            self.app.pending_sql_orders = [
                item for item in self.app.pending_sql_orders
                if self.app._item_identity_key(item) != identity
            ]
            self.app._save_pending_sql_orders()
            self.refresh()
            return

        try:
            marker.delete()
        except Exception as exc:
            messagebox.showerror("Auftrag löschen", f"Auftrag konnte nicht gelöscht werden:\n{exc}")
            return

        if marker in self.app.marker_list:
            self.app.marker_list.remove(marker)

        try:
            self.app.route_markers = [m for m in self.app.route_markers if m is not marker]
            self.app._rebuild_route_from_markers()
        except Exception:
            pass

        if getattr(self.app, "current_selected_marker", None) is marker:
            self.app.hide_info_card()
            self.app.current_selected_marker = None

        self.app.save_pins()
        try:
            self.app._refresh_all_markers()
        except Exception:
            pass
        try:
            self.app._trigger_route_metrics_recalc(force_routing=False)
        except Exception:
            pass
        self.refresh()

    def apply_filter(self):
        self._filter_job = None
        for item in self.tree.get_children():
            self.tree.delete(item)

        query = self.search_var.get().strip().lower()
        if query:
            rows = [(row, entry) for row, search_text, entry in self._all_rows if query in search_text]
        else:
            rows = [(row, entry) for row, _search_text, entry in self._all_rows]

        even_bg = getattr(self.app, "_tv_even_bg", "#ffffff")
        odd_bg = getattr(self.app, "_tv_odd_bg", "#f6f6f6")
        self.tree.tag_configure("evenrow", background=even_bg)
        self.tree.tag_configure("oddrow", background=odd_bg)
        self._row_to_entry = {}

        for index, row_data in enumerate(rows):
            d, entry = row_data
            values = tuple(d.get(col, "") for col in self.columns)
            tag = "evenrow" if index % 2 == 0 else "oddrow"
            item_id = self.tree.insert("", "end", values=values, tags=(tag,))
            self._row_to_entry[item_id] = entry

        self.info.configure(text=f"Einträge: {len(rows)}")

    def update_import_status(self):
        status_text = str(self.app.get_sql_import_status_text() or "Bereit")
        running = bool(self.app.is_sql_import_running())
        started_at = getattr(self.app, "_sql_import_started_at", None)
        if running and isinstance(started_at, datetime):
            if (datetime.now() - started_at).total_seconds() > 60 * 60:
                # Safety unlock for stale state after crashes.
                self.app._set_sql_import_status("Bereit", running=False)
                status_text = str(self.app.get_sql_import_status_text() or "Bereit")
                running = bool(self.app.is_sql_import_running())
        color = Theme.SUBTEXT
        if status_text.lower().startswith("fehler"):
            color = Theme.DANGER
        elif running:
            color = Theme.ACCENT
        self.import_status_label.configure(text=f"Importstatus: {status_text}", text_color=color)
        self.sql_import_button.configure(state="disabled" if running else "normal")

    def on_double_click(self, event):
        marker, _auftrag, _entry = self._get_selected_marker()
        if not marker:
            messagebox.showwarning(
                "Kundenkartei",
                "Dieser Auftrag ist importiert, aber noch nicht geocodiert (noch kein Pin auf der Karte).",
            )
            return

            self.app.open_customer_editor(marker)


class NonMapOrdersPage(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color=Theme.BG)
        self.app = app
        self._filter_job = None
        self._all_rows = []
        self.filter_vars = {
            label: tk.BooleanVar(value=True) for label in NON_MAP_ORDER_FILTER_OPTIONS
        }

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        topbar = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        topbar.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 12))
        topbar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(topbar, text="Nicht-Karten-Aufträge", font=_font(18, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )

        ctk.CTkButton(
            topbar,
            text="Aktualisieren",
            width=120,
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.refresh,
        ).grid(row=0, column=2, padx=12, pady=10)

        self.search_var = tk.StringVar(value="")
        search_entry = ctk.CTkEntry(
            topbar,
            textvariable=self.search_var,
            placeholder_text="Suchen (Auftragsnummer, Adresse, Produkte)...",
            width=360,
            height=36,
            corner_radius=12,
        )
        search_entry.grid(row=0, column=1, padx=(10, 8), pady=10, sticky="e")
        search_entry.bind("<KeyRelease>", self._schedule_filter, add="+")

        filter_shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        filter_shell.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        filter_shell.grid_columnconfigure(0, weight=1)
        filter_shell.grid_columnconfigure(1, weight=0)

        categories = ctk.CTkFrame(filter_shell, fg_color="transparent")
        categories.grid(row=0, column=0, padx=14, pady=12, sticky="w")

        for idx, label in enumerate(NON_MAP_ORDER_FILTER_OPTIONS):
            cb = ctk.CTkCheckBox(
                categories,
                text=label,
                variable=self.filter_vars[label],
                command=self.apply_filter,
                checkbox_width=20,
                checkbox_height=20,
                corner_radius=6,
                border_width=2,
                font=_font(12, "bold"),
                text_color=Theme.TEXT,
                fg_color=Theme.ACCENT,
                hover_color=Theme.ACCENT_HOVER,
                border_color=Theme.BORDER,
            )
            cb.grid(row=0, column=idx, padx=(0, 12), pady=2, sticky="w")

        action_buttons = ctk.CTkFrame(filter_shell, fg_color="transparent")
        action_buttons.grid(row=0, column=1, padx=14, pady=12, sticky="e")

        ctk.CTkButton(
            action_buttons,
            text="Alle",
            width=80,
            height=32,
            corner_radius=10,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self._select_all_filters,
        ).grid(row=0, column=0, padx=(0, 8), pady=2, sticky="ew")

        ctk.CTkButton(
            action_buttons,
            text="Keine",
            width=80,
            height=32,
            corner_radius=10,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self._clear_all_filters,
        ).grid(row=0, column=1, padx=(0, 0), pady=2, sticky="ew")

        self.info = ctk.CTkLabel(self, text="", anchor="w", text_color=Theme.SUBTEXT, font=_font(13))
        self.info.grid(row=2, column=0, padx=24, pady=(0, 6), sticky="w")

        table_shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        table_shell.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 20))
        table_shell.grid_columnconfigure(0, weight=1)
        table_shell.grid_rowconfigure(0, weight=1)

        self.columns = [
            "Auftragsnummer",
            "Bestelldatum",
            "Lieferart",
            "Auftragsadresse",
            "Lieferadresse",
            "Gewicht",
            "Produkte",
        ]
        self.tree = ttk.Treeview(table_shell, columns=self.columns, show="headings")
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=140, anchor="w", stretch=True)
        self.tree.column("Auftragsadresse", width=240, anchor="w", stretch=True)
        self.tree.column("Lieferadresse", width=260, anchor="w", stretch=True)
        self.tree.column("Produkte", width=420, anchor="w", stretch=True)

        vsb = ctk.CTkScrollbar(table_shell, orientation="vertical", command=self.tree.yview, **_scrollbar_kwargs())
        hsb = ctk.CTkScrollbar(table_shell, orientation="horizontal", command=self.tree.xview, **_scrollbar_kwargs())
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        vsb.grid(row=0, column=1, sticky="ns", padx=(0, 12), pady=12)
        hsb.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 12))

    def _schedule_filter(self, _event=None):
        if self._filter_job is not None:
            try:
                self.after_cancel(self._filter_job)
            except Exception:
                pass
        self._filter_job = self.after(120, self.apply_filter)

    def _select_all_filters(self):
        for var in self.filter_vars.values():
            var.set(True)
        self.apply_filter()

    def _clear_all_filters(self):
        for var in self.filter_vars.values():
            var.set(False)
        self.apply_filter()

    def refresh(self):
        self._all_rows = []
        for raw in getattr(self.app, "non_map_sql_orders", []):
            if not isinstance(raw, dict):
                continue
            row = _normalize_order_and_delivery_addresses(dict(raw))
            row["Bestelldatum"] = _display_date_string(row.get("Bestelldatum")) or str(row.get("Bestelldatum") or "")
            category = _normalize_non_map_order_category(row.get("NichtKarteKategorie") or row.get("Lieferart"))
            row["NichtKarteKategorie"] = category
            row["Lieferart"] = category or str(row.get("Lieferart") or "")
            row["Auftragsadresse"] = _format_order_address(row)
            row["Lieferadresse"] = _format_delivery_address(row)
            search_text = " ".join(str(row.get(col, "")) for col in self.columns).lower()
            self._all_rows.append((row, search_text))
        self.apply_filter()

    def apply_filter(self):
        self._filter_job = None
        for item in self.tree.get_children():
            self.tree.delete(item)

        enabled_categories = {
            label for label, var in self.filter_vars.items() if bool(var.get())
        }
        query = self.search_var.get().strip().lower()
        if not enabled_categories:
            self.info.configure(text="Einträge: 0")
            return
        rows = []
        for row, search_text in self._all_rows:
            category = str(row.get("NichtKarteKategorie") or "").strip()
            if category not in enabled_categories:
                continue
            if query and query not in search_text:
                continue
            rows.append(row)

        even_bg = getattr(self.app, "_tv_even_bg", "#ffffff")
        odd_bg = getattr(self.app, "_tv_odd_bg", "#f6f6f6")
        self.tree.tag_configure("evenrow", background=even_bg)
        self.tree.tag_configure("oddrow", background=odd_bg)

        for index, row in enumerate(rows):
            values = tuple(row.get(col, "") for col in self.columns)
            tag = "evenrow" if index % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=values, tags=(tag,))

        self.info.configure(text=f"Einträge: {len(rows)}")


class ToursPage(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color=Theme.BG)
        self.app = app
        self._all_tours = []
        self._filter_loaded_from_active_date = False

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        topbar = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        topbar.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 12))
        topbar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(topbar, text="Gespeicherte Liefertouren", font=_font(18, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )

        ctk.CTkButton(
            topbar,
            text="Aktualisieren",
            width=120,
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.refresh,
        ).grid(row=0, column=2, padx=14, pady=10)

        filter_shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        filter_shell.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 12))
        filter_shell.grid_columnconfigure(0, weight=1)
        filter_shell.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            filter_shell,
            text="Datumsfilter",
            font=_font(14, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, padx=14, pady=(12, 6), sticky="w")

        self.info = ctk.CTkLabel(filter_shell, text="", anchor="e", text_color=Theme.SUBTEXT, font=_font(13))
        self.info.grid(row=0, column=1, padx=14, pady=(12, 6), sticky="e")

        fields = ctk.CTkFrame(filter_shell, fg_color="transparent")
        fields.grid(row=1, column=0, columnspan=2, padx=14, pady=(0, 8), sticky="ew")
        fields.grid_columnconfigure((0, 1, 2), weight=1)

        self.filter_start_entry = ctk.CTkEntry(fields, height=36, corner_radius=12, placeholder_text="Von (DD.MM.YYYY)")
        self.filter_start_entry.grid(row=0, column=0, padx=(0, 8), pady=4, sticky="ew")

        self.filter_end_entry = ctk.CTkEntry(fields, height=36, corner_radius=12, placeholder_text="Bis (DD.MM.YYYY)")
        self.filter_end_entry.grid(row=0, column=1, padx=8, pady=4, sticky="ew")

        ctk.CTkButton(
            fields,
            text="Übernehmen",
            height=36,
            corner_radius=12,
            fg_color=Theme.ACCENT,
            hover_color=Theme.ACCENT_HOVER,
            command=self.apply_filters_from_inputs,
        ).grid(row=0, column=2, padx=(8, 0), pady=4, sticky="ew")

        quick = ctk.CTkFrame(filter_shell, fg_color="transparent")
        quick.grid(row=2, column=0, columnspan=2, padx=14, pady=(0, 8), sticky="ew")
        quick.grid_columnconfigure((0, 1, 2, 3), weight=1)

        ctk.CTkButton(
            quick,
            text="Heute",
            height=34,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: self.set_filter_date(format_date(datetime.now().date())),
        ).grid(row=0, column=0, padx=(0, 6), pady=4, sticky="ew")

        ctk.CTkButton(
            quick,
            text="Morgen",
            height=34,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: self.set_filter_date(format_date(datetime.now().date() + timedelta(days=1))),
        ).grid(row=0, column=1, padx=6, pady=4, sticky="ew")

        ctk.CTkButton(
            quick,
            text="Diese Woche",
            height=34,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.set_filter_this_week,
        ).grid(row=0, column=2, padx=6, pady=4, sticky="ew")

        ctk.CTkButton(
            quick,
            text="Reset",
            height=34,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.reset_filters,
        ).grid(row=0, column=3, padx=(6, 0), pady=4, sticky="ew")

        shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        shell.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 12))
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(0, weight=1)
        shell.grid_rowconfigure(1, weight=1)

        tours_shell = ctk.CTkFrame(
            shell,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        tours_shell.grid(row=0, column=0, sticky="nsew", padx=14, pady=(14, 8))
        tours_shell.grid_columnconfigure(0, weight=1)
        tours_shell.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(tours_shell, text="Touren", font=_font(14, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=12, pady=(12, 6), sticky="w"
        )

        self.cols = ["ID", "Datum", "Name", "Fahrzeug", "Mitarbeiter", "Stopps", "Totalgewicht"]
        self.tree = ttk.Treeview(tours_shell, columns=self.cols, show="headings")

        for c in self.cols:
            self.tree.heading(c, text=c)

        self.tree.column("ID", width=60, anchor="w", stretch=False)
        self.tree.column("Datum", width=110, anchor="w", stretch=False)
        self.tree.column("Name", width=180, anchor="w", stretch=True)
        self.tree.column("Fahrzeug", width=260, anchor="w", stretch=True)
        self.tree.column("Mitarbeiter", width=180, anchor="w", stretch=True)
        self.tree.column("Stopps", width=80, anchor="e", stretch=False)
        self.tree.column("Totalgewicht", width=120, anchor="e", stretch=False)

        vsb = ctk.CTkScrollbar(tours_shell, orientation="vertical", command=self.tree.yview, **_scrollbar_kwargs())
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=1, column=0, sticky="nsew", padx=(12, 0), pady=(6, 12))
        vsb.grid(row=1, column=1, sticky="ns", padx=(0, 12), pady=(6, 12))

        stops_shell = ctk.CTkFrame(
            shell,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        stops_shell.grid(row=1, column=0, sticky="nsew", padx=14, pady=(8, 14))
        stops_shell.grid_columnconfigure(0, weight=1)
        stops_shell.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            stops_shell,
            text="Stopps in ausgewählter Tour",
            font=_font(14, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, padx=12, pady=(12, 6), sticky="w")

        self.stop_cols = ["#", "Auftragsnummer", "Adresse / Name", "Gewicht"]
        self.stops_tree = ttk.Treeview(stops_shell, columns=self.stop_cols, show="headings")

        for c in self.stop_cols:
            self.stops_tree.heading(c, text=c)

        self.stops_tree.column("#", width=40, anchor="center", stretch=False)
        self.stops_tree.column("Auftragsnummer", width=140, anchor="w", stretch=False)
        self.stops_tree.column("Adresse / Name", width=360, anchor="w", stretch=True)
        self.stops_tree.column("Gewicht", width=120, anchor="e", stretch=False)

        vsb2 = ctk.CTkScrollbar(stops_shell, orientation="vertical", command=self.stops_tree.yview, **_scrollbar_kwargs())
        self.stops_tree.configure(yscrollcommand=vsb2.set)

        self.stops_tree.grid(row=1, column=0, sticky="nsew", padx=(12, 0), pady=(6, 6))
        vsb2.grid(row=1, column=1, sticky="ns", padx=(0, 12), pady=(6, 6))

        self.stops_total_label = ctk.CTkLabel(stops_shell, text="Totalgewicht: 0", font=_font(13, "bold"), text_color=Theme.TEXT)
        self.stops_total_label.grid(row=2, column=0, padx=12, pady=(0, 12), sticky="w")

        self.tour_resource_label = ctk.CTkLabel(stops_shell, text="Fahrzeug: -", font=_font(12), text_color=Theme.SUBTEXT)
        self.tour_resource_label.grid(row=3, column=0, padx=12, pady=(0, 12), sticky="w")

        btn_shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        btn_shell.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 20))
        btn_shell.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkButton(
            btn_shell,
            text="Tour auf Karte anzeigen",
            height=40,
            corner_radius=14,
            font=_font(13, "bold"),
            fg_color=Theme.ACCENT,
            hover_color=Theme.ACCENT_HOVER,
            command=self.open_selected,
        ).grid(row=0, column=0, padx=(12, 8), pady=12, sticky="ew")

        ctk.CTkButton(
            btn_shell,
            text="Tour bearbeiten",
            height=40,
            corner_radius=14,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.edit_selected,
        ).grid(row=0, column=1, padx=8, pady=12, sticky="ew")

        ctk.CTkButton(
            btn_shell,
            text="Tour löschen",
            height=40,
            corner_radius=14,
            font=_font(13, "bold"),
            fg_color=Theme.DANGER,
            hover_color=Theme.DANGER_HOVER,
            command=self.delete_selected,
        ).grid(row=0, column=2, padx=(8, 12), pady=12, sticky="ew")

        self.tree.bind("<<TreeviewSelect>>", lambda e: self._render_selected_tour_stops())
        self.tree.bind("<Double-1>", lambda e: self.open_selected())
        self.stops_tree.bind("<Double-1>", lambda e: self.on_stop_double_click())
        self.filter_start_entry.bind("<Return>", lambda e: self.apply_filters_from_inputs())
        self.filter_end_entry.bind("<Return>", lambda e: self.apply_filters_from_inputs())

        self.refresh()

    def on_stop_double_click(self):
        tour = self._get_selected_tour()
        if not tour:
            return
        sel = self.stops_tree.focus()
        if not sel:
            return
        stop = getattr(self, "_stop_row_map", {}).get(sel)
        if not stop:
            return
        marker = self.app._find_marker_for_stop(stop) if isinstance(stop, dict) else None
        if not marker:
            messagebox.showwarning("Kundenkartei", "Pin zu diesem Stopp wurde nicht gefunden.")
            return
        self.app.open_customer_editor(marker)

    def refresh(self):
        self._all_tours = self.app._load_tours()

        for item in self.tree.get_children():
            self.tree.delete(item)
        for item in self.stops_tree.get_children():
            self.stops_tree.delete(item)
        self.stops_total_label.configure(text="Totalgewicht: 0")

        tours = self._get_filtered_tours()
        even_bg = getattr(self.app, "_tv_even_bg", "#ffffff")
        odd_bg = getattr(self.app, "_tv_odd_bg", "#f6f6f6")
        self.tree.tag_configure("evenrow", background=even_bg)
        self.tree.tag_configure("oddrow", background=odd_bg)

        for index, t in enumerate(tours):
            stops = t.get("stops", []) or []
            total_w = self.app._tour_total_weight(t)
            total_str = self.app._format_weight(total_w)
            vehicle_text = self.app.format_tour_vehicle_summary(t)
            employee_text = self.app.format_employee_summary(t.get("employee_ids", []))
            values = (
                t.get("id", ""),
                _display_date_string(t.get("date", "")),
                t.get("name", ""),
                vehicle_text,
                employee_text,
                str(len(stops)),
                total_str,
            )
            tag = "evenrow" if index % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=values, tags=(tag,))

        scope = "Alle Touren"
        start_value = self.filter_start_entry.get().strip()
        end_value = self.filter_end_entry.get().strip()
        if start_value and end_value:
            scope = f"Filter: {_display_date_string(start_value)} bis {_display_date_string(end_value)}"
        elif start_value:
            scope = f"Filter: {_display_date_string(start_value)}"
        self.info.configure(text=f"{scope} | Treffer: {len(tours)}")

        children = self.tree.get_children()
        if children:
            self.tree.selection_set(children[0])
            self.tree.focus(children[0])
        self._render_selected_tour_stops()

    def on_show(self):
        active_date = getattr(self.app, "active_date", None)
        if active_date:
            self._filter_loaded_from_active_date = True
            self._set_filter_entries(active_date, active_date)
            self.refresh()

    def _set_filter_entries(self, start_value: str = "", end_value: str = ""):
        self.filter_start_entry.delete(0, "end")
        self.filter_start_entry.insert(0, str(start_value or ""))
        self.filter_end_entry.delete(0, "end")
        self.filter_end_entry.insert(0, str(end_value or ""))

    def _apply_filter_state_change(self):
        self.app.clear_active_date()
        self._filter_loaded_from_active_date = False
        self.refresh()

    def _get_filtered_tours(self):
        tours = list(self._all_tours or [])
        start_value = _normalize_date_string(self.filter_start_entry.get().strip())
        end_value = _normalize_date_string(self.filter_end_entry.get().strip())

        if start_value and end_value:
            return filter_tours_by_range(tours, start_value, end_value)
        if start_value:
            return filter_tours_by_date(tours, start_value)
        return tours

    def apply_filters_from_inputs(self):
        start_value = self.filter_start_entry.get().strip()
        end_value = self.filter_end_entry.get().strip()
        start_iso = _normalize_date_string(start_value)
        end_iso = _normalize_date_string(end_value)

        if start_value and not start_iso:
            messagebox.showwarning("Datumsfilter", "Ungültiges Startdatum. Bitte DD.MM.YYYY verwenden.")
            return
        if end_value and not end_iso:
            messagebox.showwarning("Datumsfilter", "Ungültiges Enddatum. Bitte DD.MM.YYYY verwenden.")
            return

        self._set_filter_entries(start_iso or start_value, end_iso or end_value)
        self._apply_filter_state_change()

    def set_filter_date(self, date_str):
        iso_date = _normalize_date_string(date_str)
        if not iso_date:
            return
        self._set_filter_entries(iso_date, iso_date)
        self._apply_filter_state_change()

    def set_filter_this_week(self):
        today = datetime.now().date()
        start = today - timedelta(days=today.weekday())
        end = start + timedelta(days=6)
        self._set_filter_entries(format_date(start), format_date(end))
        self._apply_filter_state_change()

    def reset_filters(self):
        self._set_filter_entries("", "")
        self._apply_filter_state_change()

    def _render_selected_tour_stops(self):
        tour = self._get_selected_tour()

        for item in self.stops_tree.get_children():
            self.stops_tree.delete(item)

        self._stop_row_map = {}
        if not tour:
            self.stops_total_label.configure(text="Totalgewicht: 0")
            self.tour_resource_label.configure(text="Fahrzeug: -")
            return

        stops = tour.get("stops", []) or []
        total = 0.0

        for i, s in enumerate(stops, start=1):
            stop_info = self.app._resolve_stop_display_info(s)
            total += stop_info["weight_value"]
            iid = f"stop_{i}"
            self._stop_row_map[iid] = s
            self.stops_tree.insert(
                "",
                "end",
                iid=iid,
                values=(i, stop_info["auftragsnummer"], stop_info["address"], stop_info["weight_text"]),
            )

        self.stops_total_label.configure(text=f"Totalgewicht: {self.app._format_weight(total)}")
        self.tour_resource_label.configure(text=f"Fahrzeug: {self.app.format_tour_vehicle_summary(tour)}")

    def _get_selected_tour(self):
        sel = self.tree.focus()
        if not sel:
            return None
        vals = self.tree.item(sel, "values")
        if not vals:
            return None
        try:
            tour_id = int(vals[0])
        except Exception:
            return None
        for t in self._all_tours:
            if t.get("id") == tour_id:
                return t
        return None

    def open_selected(self):
        tour = self._get_selected_tour()
        if not tour:
            messagebox.showwarning("Tour", "Bitte eine Tour auswählen.")
            return
        self.app.apply_tour(tour)

    def edit_selected(self):
        tour = self._get_selected_tour()
        if not tour:
            messagebox.showwarning("Tour", "Bitte eine Tour auswählen.")
            return
        self.app.open_edit_tour_dialog(tour)

    def delete_selected(self):
        tour = self._get_selected_tour()
        if not tour:
            messagebox.showwarning("Tour", "Bitte eine Tour auswählen.")
            return
        if self.app.delete_tour_record(tour.get("id"), confirm=True):
            self.refresh()


# ============================================================
# Employees
# ============================================================
class EmployeesPage(ctk.CTkFrame):
    def __init__(self, master, app):
        super().__init__(master, fg_color=Theme.BG)
        self.app = app

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        topbar = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        topbar.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 12))
        topbar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(topbar, text="Mitarbeiterverwaltung", font=_font(18, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )

        ctk.CTkButton(
            topbar,
            text="Aktualisieren",
            width=120,
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.refresh,
        ).grid(row=0, column=2, padx=(8, 8), pady=10)

        ctk.CTkButton(
            topbar,
            text="Mitarbeiter hinzufügen",
            height=36,
            corner_radius=12,
            fg_color=Theme.ACCENT,
            hover_color=Theme.ACCENT_HOVER,
            command=self.open_editor,
        ).grid(row=0, column=3, padx=(0, 14), pady=10)

        self.info = ctk.CTkLabel(self, text="", anchor="w", text_color=Theme.SUBTEXT, font=_font(13))
        self.info.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 0))

        shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=Theme.PANEL,
            border_width=1,
            border_color=Theme.BORDER,
        )
        shell.grid(row=2, column=0, sticky="nsew", padx=20, pady=(12, 20))
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(0, weight=1)

        self.scroll = ctk.CTkScrollableFrame(
            shell,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            **_scrollable_frame_kwargs(),
        )
        self.scroll.grid(row=0, column=0, sticky="nsew", padx=14, pady=14)
        self.scroll.grid_columnconfigure(0, weight=1)

        self.refresh()

    def refresh(self):
        for child in self.scroll.winfo_children():
            child.destroy()

        employees = self.app._load_employees()
        self.info.configure(text=f"Mitarbeiter: {len(employees)}")

        if not employees:
            ctk.CTkLabel(
                self.scroll,
                text="Noch keine Mitarbeiter vorhanden.",
                font=_font(13),
                text_color=Theme.SUBTEXT,
            ).grid(row=0, column=0, padx=12, pady=12, sticky="w")
            return

        for row_idx, employee in enumerate(employees):
            active = employee.get("active", True)
            card = ctk.CTkFrame(
                self.scroll,
                corner_radius=16,
                fg_color=Theme.PANEL,
                border_width=1,
                border_color=Theme.BORDER,
            )
            card.grid(row=row_idx, column=0, padx=6, pady=6, sticky="ew")
            card.grid_columnconfigure(0, weight=1)

            title_row = ctk.CTkFrame(card, fg_color="transparent")
            title_row.grid(row=0, column=0, padx=14, pady=(14, 6), sticky="ew")
            title_row.grid_columnconfigure(0, weight=1)

            name = employee.get("name", "")
            short = employee.get("short", "")
            label_text = name if not short else f"{name} ({short})"
            ctk.CTkLabel(title_row, text=label_text, font=_font(15, "bold"), text_color=Theme.TEXT).grid(
                row=0, column=0, sticky="w"
            )

            status_text = "Aktiv" if active else "Inaktiv"
            status_color = Theme.SUCCESS if active else Theme.MUTED_BTN
            ctk.CTkLabel(
                title_row,
                text=status_text,
                font=_font(11, "bold"),
                text_color=("white", "white"),
                fg_color=status_color,
                corner_radius=999,
                padx=10,
                pady=4,
            ).grid(row=0, column=1, padx=(8, 0), sticky="e")

            phone = employee.get("phone", "").strip() or "Kein Telefon hinterlegt"
            ctk.CTkLabel(card, text=phone, font=_font(12), text_color=Theme.SUBTEXT).grid(
                row=1, column=0, padx=14, pady=(0, 10), sticky="w"
            )

            btns = ctk.CTkFrame(card, fg_color="transparent")
            btns.grid(row=2, column=0, padx=14, pady=(0, 14), sticky="ew")
            btns.grid_columnconfigure((0, 1), weight=1)

            ctk.CTkButton(
                btns,
                text="Bearbeiten",
                height=36,
                corner_radius=12,
                fg_color=Theme.PANEL_2,
                hover_color=Theme.BORDER,
                text_color=Theme.TEXT,
                command=lambda value=employee: self.open_editor(value),
            ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

            ctk.CTkButton(
                btns,
                text="Löschen",
                height=36,
                corner_radius=12,
                fg_color=Theme.DANGER,
                hover_color=Theme.DANGER_HOVER,
                command=lambda value=employee: self.delete_employee(value),
            ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

    def open_editor(self, employee=None):
        employee = employee if isinstance(employee, dict) else {}

        dlg = ctk.CTkToplevel(self)
        dlg.title("Mitarbeiter")
        dlg.geometry("520x420")
        dlg.resizable(True, True)
        dlg.configure(fg_color=Theme.BG)
        dlg.attributes("-topmost", True)

        shell = ctk.CTkFrame(dlg, corner_radius=18, fg_color=Theme.PANEL, border_width=1, border_color=Theme.BORDER)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        shell.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(shell, text="Mitarbeiter", font=_font(16, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=(16, 10), sticky="w"
        )

        name_entry = ctk.CTkEntry(shell, height=36, corner_radius=12, placeholder_text="Name *")
        name_entry.grid(row=1, column=0, padx=16, pady=(0, 10), sticky="ew")
        if employee.get("name"):
            name_entry.insert(0, employee.get("name", ""))

        short_entry = ctk.CTkEntry(shell, height=36, corner_radius=12, placeholder_text="Kürzel")
        short_entry.grid(row=2, column=0, padx=16, pady=(0, 10), sticky="ew")
        if employee.get("short"):
            short_entry.insert(0, employee.get("short", ""))

        phone_entry = ctk.CTkEntry(shell, height=36, corner_radius=12, placeholder_text="Telefon")
        phone_entry.grid(row=3, column=0, padx=16, pady=(0, 10), sticky="ew")
        if employee.get("phone"):
            phone_entry.insert(0, employee.get("phone", ""))

        active_var = tk.BooleanVar(value=employee.get("active", True))
        ctk.CTkCheckBox(
            shell,
            text="Mitarbeiter ist aktiv",
            variable=active_var,
            onvalue=True,
            offvalue=False,
            text_color=Theme.TEXT,
        ).grid(row=4, column=0, padx=16, pady=(0, 12), sticky="w")

        btns = ctk.CTkFrame(shell, fg_color="transparent")
        btns.grid(row=5, column=0, padx=16, pady=(6, 16), sticky="ew")
        btns.grid_columnconfigure((0, 1), weight=1)

        def _save():
            try:
                self.app.save_employee_record({
                    "id": employee.get("id") or str(uuid4()),
                    "name": name_entry.get().strip(),
                    "short": short_entry.get().strip(),
                    "phone": phone_entry.get().strip(),
                    "active": bool(active_var.get()),
                    "created_at": employee.get("created_at") or datetime.now().replace(microsecond=0).isoformat(),
                })
            except ValueError as exc:
                messagebox.showwarning("Mitarbeiter", str(exc))
                return False
            except Exception as exc:
                messagebox.showerror("Mitarbeiter", f"Mitarbeiter konnte nicht gespeichert werden:\n{exc}")
                return

            self.refresh()
            try:
                dlg.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            btns,
            text="Abbrechen",
            height=40,
            corner_radius=14,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=dlg.destroy,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            btns,
            text="Speichern",
            height=40,
            corner_radius=14,
            font=_font(13, "bold"),
            fg_color=Theme.SUCCESS,
            hover_color=Theme.SUCCESS_HOVER,
            command=_save,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        dlg.grab_set()
        dlg.focus_force()

    def delete_employee(self, employee: dict):
        name = employee.get("name", "Mitarbeiter")
        if not messagebox.askyesno("Mitarbeiter löschen", f"{name} wirklich löschen?"):
            return False
        try:
            self.app.delete_employee_record(employee.get("id"))
        except Exception as exc:
            messagebox.showerror("Mitarbeiter", f"Mitarbeiter konnte nicht gelöscht werden:\n{exc}")
            return False
        self.refresh()


# ============================================================
# APP
# ============================================================
class ModernApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1920x1080")
        self.minsize(1500, 920)
        self.configure(fg_color=Theme.BG)

        self.depot_start_marker = None
        self.depot_end_marker = None

        self.base_dir, self.config_dir = _get_runtime_dirs()
        self._window_icon_image = None
        self.data_dir = os.path.join(self.config_dir, "data")
        self.logs_dir = os.path.join(self.config_dir, "logs")
        os.makedirs(self.config_dir, exist_ok=True)
        os.makedirs(self.data_dir, exist_ok=True)
        os.makedirs(self.logs_dir, exist_ok=True)
        self.pins_file = os.path.join(self.config_dir, "pins.json")
        self.config_file = os.path.join(self.config_dir, "settings.json")
        self.tours_file = os.path.join(self.config_dir, "tours.json")
        self.employees_file = os.path.join(self.data_dir, "employees.json")
        self.vehicles_file = os.path.join(self.data_dir, "vehicles.json")
        self.pending_orders_file = os.path.join(self.data_dir, "pending_sql_orders.json")
        self.non_map_orders_file = os.path.join(self.data_dir, "non_map_sql_orders.json")
        self.settings_manager = SettingsManager(Path(self.config_dir))
        self.sidebar_icons_dir = _resolve_runtime_asset_path(self.base_dir, "assets", "sidebar_icons")
        self.app_logo_path = _resolve_runtime_asset_path(self.base_dir, "assets", "Applogo.png")
        self.app_icon_path = _resolve_runtime_asset_path(self.base_dir, "assets", "Applogo.ico")
        self._apply_window_icon()
        set_update_log_dir(self.logs_dir)
        self._update_runtime_context = {
            "installation_type": "unknown",
            "version": "Unbekannt",
            "startup_message": "Update-Status wird geladen ...",
            "is_msix": False,
        }
        self._update_runtime_context_loading = False
        self._gps_webview_process = None

        self.geocode_cache_file = os.path.join(self.config_dir, "geocode_cache.json")
        self.geocoding_service = GeocodingService(
            self.geocode_cache_file,
            user_agent="gawela_tourenplaner_v1",
            timeout=10,
            fair_use_delay_seconds=1.1,
            retry_attempts=3,
        )
        self._route_request_job = 0
        self._search_job = 0
        self._import_job = 0
        self._pending_geocode_job = 0
        self._sql_import_running = False
        self._sql_import_status_text = "Bereit"
        self._sql_import_started_at = None
        self.pending_sql_orders = []
        self.non_map_sql_orders = []

        self.base_address = "Konstanzerstrasse 14, 8274 Tägerwilen, Schweiz"
        self.base_latlng = None
        self.base_start_active = True
        self.base_end_active = True

        # Status colors (incl. display-only "Bereits eingeplant")
        self.status_colors = {
            "Bereits eingeplant": "#94A3B8",
            "Im Lager": "#5CFF59",
            "Auf dem Weg": "#D17E0D",
            "Bestellt": "#5959FF",
            "nicht festgelegt": "#575757",
        }
        self.map_filter_status_options = ["nicht festgelegt", "Bestellt", "Auf dem Weg", "Im Lager", "Bereits eingeplant"]
        self.map_filter_selected_statuses = set(self.map_filter_status_options)
        self.map_filter_delivery_type_options = list(DELIVERY_TYPE_OPTIONS)
        self.map_filter_selected_delivery_types = set(self.map_filter_delivery_type_options)
        self.map_filter_only_current_tour = False
        self.tour_pin_color = "#8E44AD"

        self.marker_icon_size = 18
        self.marker_icons = self._build_marker_icons(self.marker_icon_size)

        self.marker_list = []
        self.search_marker = None

        self.route_points = []
        self.route_path = None
        self.route_markers = []
        self.current_route_stop_data = []
        self.current_route_segments = []
        self.current_route_summary = {}
        self.current_route_start_time = "08:00"
        self.current_route_route_mode = "car"
        self.current_route_travel_time_cache = {}
        self.current_route_distance_cache = {}
        self._route_metrics_job = 0

        self.current_route_employee_ids = []
        self.current_route_vehicle_id = None
        self.current_route_trailer_id = None
        self.current_tour_id = None
        self.active_date = None
        self.appearance_preference = "System"

        self.current_selected_marker = None
        self._selected_pin_tour = None
        self._map_filter_dialog = None
        self.map_filter_button = None

        self.sql_data_dir = DEFAULT_SQL_DATA_DIR
        self.sql_server_instance = r".\SQLEXPRESS"
        self.sql_database = infer_database_name_from_data_dir(self.sql_data_dir)

        self._last_zoom = None
        self._zoom_watch_running = False

        # Treeview zebra colors (set by _apply_treeview_style)
        self._tv_even_bg = "#FFFFFF"
        self._tv_odd_bg = "#F6F6F6"
        self.sidebar_collapsed = False
        self.sidebar_visible = True
        self.sidebar_expanded_width = 250
        self.sidebar_collapsed_width = 84
        self.sidebar_icons = self._load_sidebar_icons()
        self._employees_cache = None
        self._vehicle_data_cache = None
        self._tours_cache = None
        self._tour_data_revision = 0
        self._calendar_payload_cache = {}
        self._calendar_payload_revision = -1
        self._tour_usage_cache = {}
        self._tour_usage_cache_revision = -1
        self._depot_icon_cache = {}
        self._is_window_resizing = False
        self._last_window_size = None
        self._resize_settle_job = None
        self._pending_marker_size_job = None
        self._auto_backup_running = False
        self._route_drag_key = None
        self._route_drag_target = None
        self.quick_access_items = list(DEFAULT_QUICK_ACCESS_ITEMS)

        self._load_pending_sql_orders()
        self._load_non_map_sql_orders()

        # Layout: Sidebar + Content
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, minsize=self.sidebar_expanded_width)

        self.sidebar = ctk.CTkFrame(self, corner_radius=0, fg_color=Theme.PANEL, border_width=0, width=self.sidebar_expanded_width)
        self.sidebar.grid(row=0, column=0, sticky="nsw")
        self.sidebar.grid_propagate(False)
        self.sidebar.grid_rowconfigure(13, weight=1)

        brand = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        brand.grid(row=0, column=0, padx=16, pady=(18, 10), sticky="ew")
        brand.grid_columnconfigure(0, weight=1)
        self.brand_text = ctk.CTkFrame(brand, fg_color="transparent")
        self.brand_text.grid(row=0, column=0, sticky="w")
        self.brand_title = ctk.CTkLabel(self.brand_text, text="GAWELA", font=_font(20, "bold"), text_color=Theme.TEXT)
        self.brand_title.pack(anchor="w")
        self.brand_subtitle = ctk.CTkLabel(self.brand_text, text="Tourenplaner", font=_font(13), text_color=Theme.SUBTEXT)
        self.brand_subtitle.pack(anchor="w")
        self.btn_sidebar_toggle = ctk.CTkButton(
            brand,
            text="⟨",
            width=42,
            height=42,
            corner_radius=14,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=self.toggle_sidebar,
        )
        self.btn_sidebar_toggle.grid(row=0, column=1, padx=(8, 0), sticky="e")
        self._configure_sidebar_toggle_button()

        self.nav_menu = NavButton(
            self.sidebar,
            "Start",
            command=lambda: self.show_page("menu"),
            selected=True,
            compact_icon="⌂",
            compact_image=self.sidebar_icons.get("Start"),
        )
        self.nav_menu.grid(row=1, column=0, padx=14, pady=(10, 8), sticky="ew")

        self.nav_calendar = NavButton(
            self.sidebar,
            "Kalender",
            command=lambda: self.show_page("calendar"),
            compact_icon="◷",
            compact_image=self.sidebar_icons.get("Kalender"),
        )
        self.nav_calendar.grid(row=2, column=0, padx=14, pady=8, sticky="ew")

        self.nav_map = NavButton(
            self.sidebar,
            "Karte",
            command=lambda: self.show_page("map"),
            compact_icon="⌖",
            compact_image=self.sidebar_icons.get("Karte"),
        )
        self.nav_map.grid(row=3, column=0, padx=14, pady=8, sticky="ew")

        self.nav_gps = NavButton(
            self.sidebar,
            "GPS",
            command=lambda: self.show_page("gps"),
            compact_icon="⌕",
            compact_image=self.sidebar_icons.get("GPS"),
        )
        self.nav_gps.grid(row=4, column=0, padx=14, pady=8, sticky="ew")

        self.nav_list = NavButton(
            self.sidebar,
            "Auftragsliste",
            command=lambda: self.show_page("list"),
            compact_icon="≣",
            compact_image=self.sidebar_icons.get("Auftragsliste"),
        )
        self.nav_list.grid(row=5, column=0, padx=14, pady=8, sticky="ew")

        self.nav_nonmap = NavButton(
            self.sidebar,
            "Nicht-Karten",
            command=lambda: self.show_page("nonmap"),
            compact_icon="⋯",
        )
        self.nav_nonmap.grid(row=6, column=0, padx=14, pady=8, sticky="ew")

        self.nav_tours = NavButton(
            self.sidebar,
            "Liefertouren",
            command=lambda: self.show_page("tours"),
            compact_icon="↦",
            compact_image=self.sidebar_icons.get("Liefertouren"),
        )
        self.nav_tours.grid(row=7, column=0, padx=14, pady=8, sticky="ew")

        self.nav_employees = NavButton(
            self.sidebar,
            "Mitarbeiter",
            command=lambda: self.show_page("employees"),
            compact_icon="◉",
            compact_image=self.sidebar_icons.get("Mitarbeiter"),
        )
        self.nav_employees.grid(row=8, column=0, padx=14, pady=8, sticky="ew")

        self.nav_vehicles = NavButton(
            self.sidebar,
            "Fahrzeuge",
            command=lambda: self.show_page("vehicles"),
            compact_icon="▣",
            compact_image=self.sidebar_icons.get("Fahrzeuge"),
        )
        self.nav_vehicles.grid(row=9, column=0, padx=14, pady=8, sticky="ew")

        self.nav_settings = NavButton(
            self.sidebar,
            "Einstellungen",
            command=lambda: self.show_page("settings"),
            compact_icon="⚙",
            compact_image=self.sidebar_icons.get("Einstellungen"),
        )
        self.nav_settings.grid(row=10, column=0, padx=14, pady=8, sticky="ew")

        self.nav_update = None
        if SHOW_UPDATE_PAGE_IN_MENU:
            self.nav_update = NavButton(
                self.sidebar,
                "Updates",
                command=lambda: self.show_page("update"),
                compact_icon="↻",
            )
            self.nav_update.grid(row=11, column=0, padx=14, pady=8, sticky="ew")

        tools = ctk.CTkFrame(
            self.sidebar,
            corner_radius=16,
            fg_color=Theme.PANEL_2,
            border_width=1,
            border_color=Theme.BORDER,
        )
        tools.grid(row=12, column=0, padx=14, pady=(10, 12), sticky="ew")
        tools.grid_columnconfigure(0, weight=1)
        self.sidebar_tools = tools

        self.tools_title = ctk.CTkLabel(tools, text="Schnellzugriff", font=_font(13, "bold"), text_color=Theme.TEXT)
        self.tools_title.grid(
            row=0, column=0, padx=12, pady=(12, 8), sticky="w"
        )
        self.quick_access_container = ctk.CTkFrame(tools, fg_color="transparent")
        self.quick_access_container.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="ew")
        self.quick_access_container.grid_columnconfigure(0, weight=1)
        self.quick_access_buttons = []
        self.refresh_quick_access_tools()

        self.sidebar_footer = ctk.CTkLabel(self.sidebar, text="© GAWELA", font=_font(12), text_color=Theme.SUBTEXT)
        self.sidebar_footer.grid(row=14, column=0, padx=14, pady=(0, 14), sticky="w")

        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.grid(row=0, column=1, sticky="nsew")
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.pages = {
            "menu": StartMenuPage(self.container, self),
            "calendar": CalendarPage(self.container, self),
            "map": MapPage(self.container, self),
            "gps": GPSPage(self.container, self),
            "list": OrdersListPage(self.container, self),
            "nonmap": NonMapOrdersPage(self.container, self),
            "tours": ToursPage(self.container, self),
            "employees": EmployeesPage(self.container, self),
            "vehicles": VehiclesPage(self.container, self, Theme, _font),
            "settings": SettingsPage(self.container, self, Theme, _font),
            "update": UpdatePage(self.container, self, Theme, _font),
        }
        self._page_grid_kwargs = {"row": 0, "column": 0, "sticky": "nsew"}
        for page in self.pages.values():
            page.grid(**self._page_grid_kwargs)
            page.grid_remove()

        self.update_route_employee_summary()
        self.update_route_resource_summary()
        self._apply_treeview_style()
        self.load_config()
        self.load_pins()
        self.import_from_sql(silent=True)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.bind("<Configure>", self._on_window_configure, add="+")
        self.show_page("menu")
        self.refresh_update_runtime_context()
        self.start_zoom_watch()
        self.after(2500, self._check_auto_backup_due)

    def _apply_window_icon(self):
        if os.path.exists(self.app_logo_path):
            try:
                self._window_icon_image = tk.PhotoImage(file=self.app_logo_path)
                self.iconphoto(True, self._window_icon_image)
            except Exception:
                logger.exception("PNG app icon could not be applied.")
        if os.path.exists(self.app_icon_path):
            try:
                self.iconbitmap(self.app_icon_path)
            except Exception:
                logger.exception("ICO app icon could not be applied.")

    def _on_window_configure(self, event=None):
        if event is None or getattr(event, "widget", None) is not self:
            return
        try:
            size = (int(event.width), int(event.height))
        except Exception:
            return
        if self._last_window_size == size:
            return
        self._last_window_size = size
        self._is_window_resizing = True
        if self._resize_settle_job is not None:
            try:
                self.after_cancel(self._resize_settle_job)
            except Exception:
                pass
        self._resize_settle_job = self.after(220, self._on_window_resize_settled)

    def _on_window_resize_settled(self):
        self._resize_settle_job = None
        self._is_window_resizing = False
        if getattr(self, "current_page", "") == "map" and self._last_zoom is not None:
            self._schedule_marker_size_update(self._last_zoom, delay_ms=120)

    def _get_gps_helper_command(self) -> list[str]:
        if getattr(sys, "frozen", False):
            return [sys.executable, "--gps-webview", GPS_TRACKING_URL]

        project_pythonw = os.path.join(self.base_dir, ".venv313", "Scripts", "pythonw.exe")
        project_python = os.path.join(self.base_dir, ".venv313", "Scripts", "python.exe")
        if os.path.exists(project_pythonw):
            return [project_pythonw, os.path.abspath(__file__), "--gps-webview", GPS_TRACKING_URL]
        if os.path.exists(project_python):
            return [project_python, os.path.abspath(__file__), "--gps-webview", GPS_TRACKING_URL]

        current_dir = os.path.dirname(sys.executable)
        current_pythonw = os.path.join(current_dir, "pythonw.exe")
        if os.path.exists(current_pythonw):
            return [current_pythonw, os.path.abspath(__file__), "--gps-webview", GPS_TRACKING_URL]
        return [sys.executable, os.path.abspath(__file__), "--gps-webview", GPS_TRACKING_URL]

    def get_gps_embed_helper_command(self, parent_hwnd: int, width: int, height: int, url: str) -> list[str]:
        if getattr(sys, "frozen", False):
            return [sys.executable, "--gps-embed", str(parent_hwnd), str(width), str(height), str(url or GPS_TRACKING_URL)]

        project_pythonw = os.path.join(self.base_dir, ".venv313", "Scripts", "pythonw.exe")
        project_python = os.path.join(self.base_dir, ".venv313", "Scripts", "python.exe")
        if os.path.exists(project_pythonw):
            return [
                project_pythonw,
                os.path.abspath(__file__),
                "--gps-embed",
                str(parent_hwnd),
                str(width),
                str(height),
                str(url or GPS_TRACKING_URL),
            ]
        if os.path.exists(project_python):
            return [
                project_python,
                os.path.abspath(__file__),
                "--gps-embed",
                str(parent_hwnd),
                str(width),
                str(height),
                str(url or GPS_TRACKING_URL),
            ]
        return [
            sys.executable,
            os.path.abspath(__file__),
            "--gps-embed",
            str(parent_hwnd),
            str(width),
            str(height),
            str(url or GPS_TRACKING_URL),
        ]

    def _is_gps_webview_running(self) -> bool:
        process = getattr(self, "_gps_webview_process", None)
        return bool(process and process.poll() is None)

    def get_gps_runtime_status(self) -> str:
        runtime_path = _resolve_webview2_runtime_path(self.base_dir)
        if runtime_path:
            if _is_runtime_inside_bundle(self.base_dir, runtime_path):
                runtime_hint = f"Gebundene Fixed-Version-Runtime aktiv: {runtime_path}"
            else:
                runtime_hint = f"Systeminstallierte WebView2-Runtime aktiv: {runtime_path}"
        else:
            runtime_hint = "Es wurde keine WebView2-Runtime gefunden."
        helper_command = self._get_gps_helper_command()
        helper_runtime = helper_command[0] if helper_command else "Nicht gefunden"
        if self._is_gps_webview_running():
            return f"Native WebView2 aktiv. {runtime_hint} Helper: {helper_runtime}"
        return f"Bereit für native WebView2. {runtime_hint} Helper: {helper_runtime}"

    def terminate_gps_native_window(self, *, timeout: float = 5.0) -> bool:
        process = getattr(self, "_gps_webview_process", None)
        success = _terminate_process_gracefully(process, timeout=timeout)
        if success:
            self._gps_webview_process = None
        return success

    def launch_gps_native_window(self, *, force: bool = False) -> bool:
        if self._is_gps_webview_running():
            if not force:
                return True
            self.terminate_gps_native_window()

        command = self._get_gps_helper_command()
        startupinfo = None
        creationflags = 0
        if os.name == "nt":
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

        env, runtime_path = _apply_webview2_runtime_environment(self.base_dir)
        if runtime_path:
            env["GAWELA_WEBVIEW2_RUNTIME_PATH"] = runtime_path

        self._gps_webview_process = subprocess.Popen(
            command,
            cwd=self.base_dir,
            env=env,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
        return True

    # ---------- Darkmode Toggle ----------
    def toggle_darkmode(self):
        """Toggle between Light and Dark, refresh ttk styles + views."""
        try:
            mode = ctk.get_appearance_mode()
            self.set_appearance_preference("Light" if mode == "Dark" else "Dark")
        except Exception:
            return

    def get_appearance_preference(self) -> str:
        return self.appearance_preference or "System"

    def set_appearance_preference(self, mode: str, persist: bool = True):
        normalized = str(mode or "System").strip().title()
        if normalized not in {"System", "Light", "Dark"}:
            normalized = "System"
        try:
            ctk.set_appearance_mode(normalized)
        except Exception:
            return
        self.appearance_preference = normalized
        if persist:
            self.save_config()

        self._apply_treeview_style()

        for page_name in ("list", "tours", "menu", "map", "vehicles", "settings", "update"):
            try:
                page = self.pages.get(page_name)
                if page is None:
                    continue
                if hasattr(page, "refresh"):
                    page.refresh()
                if hasattr(page, "on_theme_changed"):
                    page.on_theme_changed()
            except Exception:
                pass

    # ---------- Navigation ----------
    def _set_nav_selected(self, key: str):
        self.nav_menu.set_selected(key == "menu")
        self.nav_calendar.set_selected(key == "calendar")
        self.nav_map.set_selected(key == "map")
        self.nav_gps.set_selected(key == "gps")
        self.nav_list.set_selected(key == "list")
        self.nav_nonmap.set_selected(key == "nonmap")
        self.nav_tours.set_selected(key == "tours")
        self.nav_employees.set_selected(key == "employees")
        self.nav_vehicles.set_selected(key == "vehicles")
        self.nav_settings.set_selected(key == "settings")
        if self.nav_update is not None:
            self.nav_update.set_selected(key == "update")

    def _configure_sidebar_toggle_button(self):
        if self.sidebar_collapsed:
            self.btn_sidebar_toggle.configure(
                text="⟩",
                width=30,
                height=30,
                corner_radius=12,
            )
        else:
            self.btn_sidebar_toggle.configure(
                text="⟨",
                width=42,
                height=42,
                corner_radius=14,
            )

    def _set_sidebar_toggle_text(self, collapsed: bool):
        self.btn_sidebar_toggle.configure(text="⟩" if collapsed else "⟨")

    def _sidebar_nav_buttons(self):
        buttons = [
            self.nav_menu,
            self.nav_calendar,
            self.nav_map,
            self.nav_gps,
            self.nav_list,
            self.nav_nonmap,
            self.nav_tours,
            self.nav_employees,
            self.nav_vehicles,
            self.nav_settings,
        ]
        if self.nav_update is not None:
            buttons.append(self.nav_update)
        return tuple(buttons)

    def _apply_sidebar_layout_state(self, collapsed: bool):
        if collapsed:
            self.brand_text.grid_remove()
            self.sidebar_footer.grid_remove()
            self.sidebar_tools.grid_remove()
        else:
            self.brand_text.grid()
            self.sidebar_footer.grid()
            self.sidebar_tools.grid()

        for button in self._sidebar_nav_buttons():
            button.set_compact(collapsed)
            button.grid_configure(padx=10 if collapsed else 14, pady=6 if collapsed else 8)

        if not collapsed:
            self.nav_menu.grid_configure(pady=(10, 8))

    def _set_sidebar_width(self, width: int):
        width = max(self.sidebar_collapsed_width, min(self.sidebar_expanded_width, int(width)))
        self.grid_columnconfigure(0, minsize=width)
        self.sidebar.configure(width=width)

    def _get_sidebar_target_width(self) -> int:
        return self.sidebar_collapsed_width if self.sidebar_collapsed else self.sidebar_expanded_width

    def _set_sidebar_visible(self, visible: bool):
        visible = bool(visible)
        if visible == self.sidebar_visible:
            if visible:
                self._set_sidebar_width(self._get_sidebar_target_width())
                self.container.grid_configure(column=1, columnspan=1)
            return

        self.sidebar_visible = visible
        if visible:
            self.sidebar.grid()
            self._set_sidebar_width(self._get_sidebar_target_width())
            self.container.grid_configure(column=1, columnspan=1)
        else:
            self.sidebar.grid_remove()
            self.grid_columnconfigure(0, minsize=0)
            self.container.grid_configure(column=0, columnspan=2)

    def toggle_sidebar(self):
        self.sidebar_collapsed = not self.sidebar_collapsed
        target_width = self._get_sidebar_target_width()
        self._set_sidebar_toggle_text(self.sidebar_collapsed)
        self._apply_sidebar_layout_state(self.sidebar_collapsed)
        if self.sidebar_visible:
            self._set_sidebar_width(target_width)
        self._configure_sidebar_toggle_button()

    def get_quick_access_options(self) -> list[tuple[str, str]]:
        return [
            ("", "Kein Eintrag"),
            ("action:export_route", "Route exportieren"),
            ("action:import_sql", "SQL importieren"),
            ("action:select_sql_dir", "SQL-Datenordner wählen"),
            ("page:menu", "Start"),
            ("page:calendar", "Kalender"),
            ("page:map", "Karte"),
            ("page:gps", "GPS"),
            ("page:list", "Auftragsliste"),
            ("page:nonmap", "Nicht-Karten-Aufträge"),
            ("page:tours", "Liefertouren"),
            ("page:employees", "Mitarbeiter"),
            ("page:vehicles", "Fahrzeuge"),
            ("page:settings", "Einstellungen"),
        ] + ([("page:update", "Updates")] if SHOW_UPDATE_PAGE_IN_MENU else [])

    def get_quick_access_option_map(self) -> dict:
        return {key: label for key, label in self.get_quick_access_options()}

    def normalize_quick_access_items(self, items) -> list[str]:
        legacy_map = {
            "action:import_folder": "action:import_sql",
            "action:select_xml": "action:select_sql_dir",
        }
        valid_ids = {key for key, _label in self.get_quick_access_options()}
        normalized = []
        for value in items if isinstance(items, list) else []:
            key = legacy_map.get(str(value or "").strip(), str(value or "").strip())
            if key not in valid_ids:
                key = ""
            if key and key in normalized:
                key = ""
            normalized.append(key)
        while len(normalized) < len(DEFAULT_QUICK_ACCESS_ITEMS):
            normalized.append("")
        return normalized[: len(DEFAULT_QUICK_ACCESS_ITEMS)]

    def refresh_quick_access_tools(self):
        container = getattr(self, "quick_access_container", None)
        if container is None:
            return

        for button in getattr(self, "quick_access_buttons", []):
            try:
                button.destroy()
            except Exception:
                pass
        self.quick_access_buttons = []

        row = 0
        for item_id in self.normalize_quick_access_items(getattr(self, "quick_access_items", [])):
            button = self._create_quick_access_button(container, item_id)
            if button is None:
                continue
            button.grid(row=row, column=0, pady=6 if row else 0, sticky="ew")
            self.quick_access_buttons.append(button)
            row += 1

    def _create_quick_access_button(self, master, item_id: str):
        item_id = str(item_id or "").strip()
        if not item_id:
            return None

        config = {
            "height": 36,
            "corner_radius": 12,
            "font": _font(13, "bold"),
        }
        if item_id == "action:select_sql_dir":
            config.update(
                text="SQL-Datenordner wählen",
                fg_color=Theme.PANEL,
                hover_color=Theme.BORDER,
                text_color=Theme.TEXT,
                command=self.select_sql_data_dir,
            )
        elif item_id == "action:import_sql":
            config.update(
                text="SQL importieren",
                fg_color=Theme.PANEL,
                hover_color=Theme.BORDER,
                text_color=Theme.TEXT,
                command=self.import_from_sql,
            )
        elif item_id == "action:export_route":
            config.update(
                text="Route exportieren",
                fg_color=Theme.SUCCESS,
                hover_color=Theme.SUCCESS_HOVER,
                text_color=("white", "white"),
                command=self.export_route,
            )
        elif item_id.startswith("page:"):
            page_name = item_id.split(":", 1)[1]
            label = self.get_quick_access_option_map().get(item_id)
            if not label:
                return None
            config.update(
                text=label,
                fg_color=Theme.PANEL,
                hover_color=Theme.BORDER,
                text_color=Theme.TEXT,
                command=lambda name=page_name: self.show_page(name),
            )
        else:
            return None

        return ctk.CTkButton(master, **config)

    def set_active_date(self, date_str):
        normalized = _normalize_date_string(date_str)
        self.active_date = normalized or None

    def clear_active_date(self):
        self.active_date = None

    def show_page(self, name: str):
        page = self.pages.get(name)
        if not page:
            messagebox.showerror("Navigation", f"Seite '{name}' ist nicht registriert.")
            return
        previous_page = self.pages.get(getattr(self, "current_page", ""))
        if previous_page is not None and previous_page is not page and hasattr(previous_page, "on_hide"):
            try:
                previous_page.on_hide()
            except Exception:
                pass
        if previous_page is not None and previous_page is not page:
            try:
                previous_page.grid_remove()
            except Exception:
                pass
        try:
            if str(page.winfo_manager()) != "grid":
                page.grid(**self._page_grid_kwargs)
        except Exception:
            page.grid(**self._page_grid_kwargs)
        self._set_sidebar_visible(name != "menu")
        self.current_page = name
        page.tkraise()
        self._set_nav_selected(name)
        if name in ("menu", "calendar", "gps", "list", "nonmap", "tours", "employees", "vehicles", "settings", "update") and hasattr(page, "refresh"):
            page.refresh()
        if hasattr(page, "on_show"):
            try:
                page.on_show()
            except Exception:
                pass

    def refresh_update_runtime_context(self):
        if self._update_runtime_context_loading:
            return
        self._update_runtime_context_loading = True

        def _worker():
            try:
                context = get_runtime_update_context()
            except Exception:
                logger.exception("Update runtime context could not be loaded.")
                context = {
                    "installation_type": "unknown",
                    "version": "Unbekannt",
                    "is_msix": False,
                }
            context["startup_message"] = self._build_startup_update_message(context)
            self.after(0, lambda: self._apply_update_runtime_context(context))

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_update_runtime_context(self, context: dict):
        self._update_runtime_context_loading = False
        self._update_runtime_context = dict(context or {})
        try:
            menu_page = self.pages.get("menu")
            if menu_page and hasattr(menu_page, "refresh"):
                menu_page.refresh()
        except Exception:
            logger.exception("Start menu could not be refreshed after update context load.")

    def _build_startup_update_message(self, context: dict) -> str:
        version = str(context.get("version") or "Unbekannt")
        if context.get("installation_type") == "msix":
            return f"Version {version}. Updates werden automatisch über App Installer geprüft, wenn die Installation über .appinstaller erfolgt ist."
        return f"Version {version}. Für Auto-Updates bitte die App über die .appinstaller-Datei als MSIX installieren."

    def get_startup_update_message(self) -> str:
        return str(self._update_runtime_context.get("startup_message") or "Update-Status wird geladen ...")

    # ---------- Treeview Style ----------
    def _apply_treeview_style(self):
        """ttk.Treeview sauber für Light/Dark stylen."""
        try:
            style = ttk.Style()
            try:
                style.theme_use("clam")
            except Exception:
                pass

            mode = ctk.get_appearance_mode()  # "Light" / "Dark"
            if mode == "Dark":
                panel = "#151822"
                header_bg = "#1F2430"
                fg = "#E7EAF3"
                header_fg = "#E7EAF3"
                sel_bg = "#24314A"
                sel_fg = "#E7EAF3"
                self._tv_even_bg = panel
                self._tv_odd_bg = "#10131A"
            else:
                panel = "#FFFFFF"
                header_bg = "#EEF2FF"
                fg = "#111827"
                header_fg = "#111827"
                sel_bg = "#DBEAFE"
                sel_fg = "#111827"
                self._tv_even_bg = panel
                self._tv_odd_bg = "#F6F6F6"

            style.configure(
                "Treeview",
                rowheight=28,
                borderwidth=0,
                relief="flat",
                background=panel,
                fieldbackground=panel,
                foreground=fg,
            )
            style.configure(
                "Treeview.Heading",
                borderwidth=0,
                relief="flat",
                background=header_bg,
                foreground=header_fg,
                font=("Segoe UI", 10, "bold"),
            )
            style.configure(
                "Vertical.TScrollbar",
                background=Theme.resolve(Theme.SCROLLBAR_BUTTON),
                troughcolor=panel,
                bordercolor=panel,
                arrowcolor=Theme.resolve(Theme.TEXT),
                darkcolor=Theme.resolve(Theme.SCROLLBAR_BUTTON),
                lightcolor=Theme.resolve(Theme.SCROLLBAR_BUTTON),
                gripcount=0,
            )
            style.map(
                "Vertical.TScrollbar",
                background=[("active", Theme.resolve(Theme.SCROLLBAR_HOVER))],
            )
            style.configure(
                "Horizontal.TScrollbar",
                background=Theme.resolve(Theme.SCROLLBAR_BUTTON),
                troughcolor=panel,
                bordercolor=panel,
                arrowcolor=Theme.resolve(Theme.TEXT),
                darkcolor=Theme.resolve(Theme.SCROLLBAR_BUTTON),
                lightcolor=Theme.resolve(Theme.SCROLLBAR_BUTTON),
                gripcount=0,
            )
            style.map(
                "Horizontal.TScrollbar",
                background=[("active", Theme.resolve(Theme.SCROLLBAR_HOVER))],
            )
            combo_arrow = panel
            combo_select = Theme.resolve(Theme.SELECTION)
            style.configure(
                "Modern.TCombobox",
                fieldbackground=panel,
                background=combo_arrow,
                foreground=fg,
                arrowcolor=header_fg,
                bordercolor=panel,
                lightcolor=combo_arrow,
                darkcolor=combo_arrow,
                insertcolor=fg,
                relief="flat",
                padding=4,
                arrowsize=12,
            )
            style.map(
                "Modern.TCombobox",
                fieldbackground=[("readonly", panel), ("focus", panel), ("active", panel)],
                background=[("active", header_bg), ("readonly", combo_arrow)],
                foreground=[("readonly", fg), ("focus", fg)],
                arrowcolor=[("active", Theme.resolve(Theme.ACCENT)), ("readonly", Theme.resolve(Theme.ACCENT))],
                selectbackground=[("focus", combo_select)],
                selectforeground=[("focus", fg)],
            )
            style.map("Treeview", background=[("selected", sel_bg)], foreground=[("selected", sel_fg)])
            style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        except Exception:
            self._tv_even_bg = "#FFFFFF"
            self._tv_odd_bg = "#F6F6F6"

    # ---------- Weight helpers ----------
    def _parse_weight(self, w) -> float:
        if w is None:
            return 0.0
        s = str(w).strip()
        if not s or s.upper() == "N/A":
            return 0.0
        s = s.lower().replace("kg", "").strip()
        s = s.replace(" ", "").replace(",", ".")
        num = ""
        for ch in s:
            if ch.isdigit() or ch == ".":
                num += ch
            else:
                if num:
                    break
        try:
            return float(num) if num else 0.0
        except Exception:
            return 0.0

    def _format_weight(self, w: float) -> str:
        if w is None:
            return "0 kg"
        try:
            if abs(w - round(w)) < 1e-9:
                value = f"{int(round(w))}"
            else:
                value = f"{w:.2f}".rstrip("0").rstrip(".")
            return f"{value} kg"
        except Exception:
            return "0 kg"

    def _marker_weight_value(self, marker) -> float:
        d = getattr(marker, "data", {}) or {}
        return self._parse_weight(d.get("Gewicht", ""))

    def _marker_weight_text(self, marker) -> str:
        d = getattr(marker, "data", {}) or {}
        raw = d.get("Gewicht", "")
        val = self._parse_weight(raw)
        return self._format_weight(val)

    def _tour_total_weight(self, tour: dict) -> float:
        total = 0.0
        if not tour or not isinstance(tour, dict):
            return total
        stops = tour.get("stops", []) or []
        for s in stops:
            if isinstance(s, dict):
                if s.get("weight") is not None:
                    total += self._parse_weight(s.get("weight"))
                else:
                    m = self._find_marker_for_stop(s)
                    if m:
                        total += self._marker_weight_value(m)
        return total

    def _find_marker_for_stop(self, s: dict):
        auftrag = (s.get("auftragsnummer") or "").strip()
        lat = s.get("lat")
        lng = s.get("lng")

        if auftrag:
            for m in self.marker_list:
                if getattr(m, "is_system", False):
                    continue
                if str(getattr(m, "auftragsnummer", "")).strip() == auftrag:
                    return m

        if lat is not None and lng is not None:
            for m in self.marker_list:
                if getattr(m, "is_system", False):
                    continue
                try:
                    if abs(m.position[0] - float(lat)) < 1e-7 and abs(m.position[1] - float(lng)) < 1e-7:
                        return m
                except Exception:
                    pass

        return None

    def _resolve_stop_display_info(self, s: dict) -> dict:
        auftrag = ""
        address = ""
        weight_text = "0"
        weight_value = 0.0

        if not isinstance(s, dict):
            return {"auftragsnummer": auftrag, "address": address, "weight_text": weight_text, "weight_value": weight_value}

        auftrag = (s.get("auftragsnummer") or "").strip()
        m = self._find_marker_for_stop(s)

        if m:
            d = getattr(m, "data", {}) or {}
            name = str(d.get("Name", "")).strip()
            street = str(d.get("Strasse", "")).strip()
            plz = str(d.get("PLZ", "")).strip()
            ort = str(d.get("Ort", "")).strip()
            line2 = ", ".join([p for p in [street, " ".join([plz, ort]).strip()] if p]).strip()
            if name and line2:
                address = f"{name} / {line2}"
            else:
                address = name or line2 or ""
            weight_value = self._marker_weight_value(m)
            weight_text = self._format_weight(weight_value)
            if not auftrag:
                auftrag = str(getattr(m, "auftragsnummer", "")).strip()
        else:
            if s.get("weight") is not None:
                weight_value = self._parse_weight(s.get("weight"))
                weight_text = self._format_weight(weight_value)
            address = ""

        return {"auftragsnummer": auftrag, "address": address, "weight_text": weight_text, "weight_value": weight_value}

    def _stop_key(self, stop: dict):
        if not isinstance(stop, dict):
            return None
        auftrag = str(stop.get("auftragsnummer") or "").strip()
        if auftrag:
            return ("auftrag", auftrag)
        lat = stop.get("lat")
        lon = stop.get("lon", stop.get("lng"))
        if lat is not None and lon is not None:
            try:
                return ("coord", round(float(lat), 7), round(float(lon), 7))
            except Exception:
                return None
        stop_id = str(stop.get("id") or "").strip()
        if stop_id:
            return ("stop", stop_id)
        return None

    def _format_time_window(self, stop: dict) -> str:
        start = str((stop or {}).get("time_window_start") or "").strip()
        end = str((stop or {}).get("time_window_end") or "").strip()
        if start and end:
            return f"{start}-{end}"
        if start:
            return f"ab {start}"
        if end:
            return f"bis {end}"
        return "—"

    def _make_default_stop_from_marker(self, marker, existing: dict | None = None, order: int = 0):
        existing = dict(existing or {})
        d = getattr(marker, "data", {}) or {}
        lat = None
        lon = None
        try:
            lat, lon = float(marker.position[0]), float(marker.position[1])
        except Exception:
            pass

        name = str(existing.get("name") or d.get("Name") or getattr(marker, "auftragsnummer", "") or "").strip()
        street = str(d.get("Strasse", "")).strip()
        plz = str(d.get("PLZ", "")).strip()
        ort = str(d.get("Ort", "")).strip()
        address = str(existing.get("address") or "").strip()
        if not address:
            address = ", ".join([value for value in [street, " ".join([plz, ort]).strip()] if value])

        stop_id = str(existing.get("id") or "").strip()
        if not stop_id:
            marker_key = self._marker_key(marker)
            if marker_key and marker_key[0] == "auftrag":
                stop_id = f"auftrag:{marker_key[1]}"
            elif marker_key and marker_key[0] == "coord":
                stop_id = f"coord:{marker_key[1]}:{marker_key[2]}"
            else:
                stop_id = str(uuid4())

        service_minutes = existing.get("service_minutes", 0)
        try:
            service_minutes = int(service_minutes or 0)
        except Exception:
            service_minutes = 0

        stop = dict(existing)
        stop["id"] = stop_id
        stop["name"] = name
        stop["address"] = address
        stop["lat"] = lat
        stop["lon"] = lon
        stop["lng"] = lon
        stop["order"] = order
        stop["auftragsnummer"] = str(existing.get("auftragsnummer") or getattr(marker, "auftragsnummer", "") or "").strip() or None
        stop["weight"] = existing.get("weight", d.get("Gewicht", ""))
        stop["time_window_start"] = str(existing.get("time_window_start") or "").strip()
        stop["time_window_end"] = str(existing.get("time_window_end") or "").strip()
        stop["service_minutes"] = service_minutes
        stop["planned_arrival"] = str(existing.get("planned_arrival") or "").strip()
        stop["planned_departure"] = str(existing.get("planned_departure") or "").strip()
        stop["schedule_conflict"] = bool(existing.get("schedule_conflict", False))
        stop["schedule_conflict_text"] = str(existing.get("schedule_conflict_text") or "").strip()
        stop["wait_minutes"] = int(existing.get("wait_minutes") or 0)
        return stop

    def _sync_current_route_stop_data_from_markers(self, preferred_stops=None):
        preferred_map = {}
        for stop in preferred_stops or []:
            key = self._stop_key(stop)
            if key is not None:
                preferred_map[key] = dict(stop)

        current_map = {}
        for stop in getattr(self, "current_route_stop_data", []):
            key = self._stop_key(stop)
            if key is not None:
                current_map[key] = dict(stop)

        ordered = []
        order = 1
        for marker in getattr(self, "route_markers", []):
            if not marker or self._is_depot_marker(marker):
                continue
            marker_key = self._marker_key(marker)
            base = preferred_map.get(marker_key) or current_map.get(marker_key)
            ordered.append(self._make_default_stop_from_marker(marker, existing=base, order=order))
            order += 1

        self.current_route_stop_data = ordered

    def _build_route_nodes(self):
        nodes = []
        if self.depot_start_marker:
            nodes.append({
                "id": "depot_start",
                "name": "Start (Depot)",
                "address": self.base_address,
                "lat": self.depot_start_marker.position[0],
                "lon": self.depot_start_marker.position[1],
            })
        nodes.extend([dict(stop) for stop in getattr(self, "current_route_stop_data", [])])
        if self.depot_end_marker:
            nodes.append({
                "id": "depot_end",
                "name": "Ende (Depot)",
                "address": self.base_address,
                "lat": self.depot_end_marker.position[0],
                "lon": self.depot_end_marker.position[1],
            })
        return nodes

    def _apply_schedule_results(self, segment_results: list):
        segments = list(segment_results or [])
        stops = [dict(stop) for stop in getattr(self, "current_route_stop_data", [])]
        stop_segment_minutes = [segment.get("travel_minutes") for segment in segments[: len(stops)]]
        schedule = compute_schedule(stops, stop_segment_minutes, self.current_route_start_time)
        self.current_route_stop_data = schedule.get("stops", stops)
        stop_by_id = {str(stop.get("id") or ""): stop for stop in self.current_route_stop_data}
        end_time = str(schedule.get("end_time") or "")

        detailed_segments = []
        nodes = self._build_route_nodes()
        for index, segment in enumerate(segments):
            from_node = nodes[index] if index < len(nodes) else {}
            to_node = nodes[index + 1] if index + 1 < len(nodes) else {}
            detail = dict(segment)
            detail["from_name"] = from_node.get("name") or ("Depot" if index == 0 else f"Stopp {index}")
            detail["to_name"] = to_node.get("name") or f"Stopp {index + 1}"
            from_stop = stop_by_id.get(str(from_node.get("id") or ""))
            to_stop = stop_by_id.get(str(to_node.get("id") or ""))

            departure = ""
            arrival = ""
            wait_minutes = 0
            if str(from_node.get("id") or "") == "depot_start":
                departure = self.current_route_start_time
            elif from_stop:
                departure = str(from_stop.get("planned_departure") or "")

            if str(to_node.get("id") or "") == "depot_end":
                arrival = end_time
            elif to_stop:
                arrival = str(to_stop.get("planned_arrival") or "")
                wait_minutes = int(to_stop.get("wait_minutes") or 0)

            detail["departure"] = departure
            detail["arrival"] = arrival
            detail["wait_minutes"] = wait_minutes
            detailed_segments.append(detail)

        self.current_route_segments = detailed_segments
        self.current_route_summary = {
            "total_travel_minutes": schedule.get("total_travel_minutes", 0),
            "total_service_minutes": schedule.get("total_service_minutes", 0),
            "total_wait_minutes": schedule.get("total_wait_minutes", 0),
            "end_time": schedule.get("end_time", ""),
            "has_conflicts": schedule.get("has_conflicts", False),
        }

    def _update_route_metrics_ui(self, status_text: str = ""):
        self.refresh_route_panel()
        segment_frame = getattr(self, "route_segment_list", None)
        if segment_frame is not None:
            for child in segment_frame.winfo_children():
                child.destroy()

            if not self.current_route_segments:
                ctk.CTkLabel(
                    segment_frame,
                    text=status_text or "Noch keine Fahrzeiten verfügbar.",
                    font=_font(12),
                    text_color=Theme.SUBTEXT,
                ).grid(row=0, column=0, padx=10, pady=10, sticky="w")
            else:
                for row_idx, segment in enumerate(self.current_route_segments):
                    travel_minutes = segment.get("travel_minutes")
                    distance_km = segment.get("distance_km")
                    travel_text = "—" if travel_minutes is None else f"{travel_minutes} min"
                    distance_text = "" if distance_km is None else f" · {distance_km:.1f} km"
                    if int(segment.get("wait_minutes") or 0) > 0:
                        distance_text += f" · Warten {int(segment.get('wait_minutes') or 0)} min"
                    timing = ""
                    if segment.get("departure") or segment.get("arrival"):
                        timing = f"{segment.get('departure') or '—'} → {segment.get('arrival') or '—'}"

                    row = ctk.CTkFrame(
                        segment_frame,
                        corner_radius=12,
                        fg_color=Theme.PANEL,
                        border_width=1,
                        border_color=Theme.BORDER,
                    )
                    row.grid(row=row_idx, column=0, padx=4, pady=4, sticky="ew")
                    row.grid_columnconfigure(0, weight=1)

                    ctk.CTkLabel(
                        row,
                        text=f"{segment.get('from_name', '')} → {segment.get('to_name', '')}",
                        font=_font(12, "bold"),
                        text_color=Theme.TEXT,
                    ).grid(row=0, column=0, padx=10, pady=(8, 2), sticky="w")

                    ctk.CTkLabel(
                        row,
                        text=f"{travel_text}{distance_text}",
                        font=_font(12),
                        text_color=Theme.SUBTEXT,
                    ).grid(row=1, column=0, padx=10, pady=(0, 2), sticky="w")

                    if timing:
                        ctk.CTkLabel(
                            row,
                            text=timing,
                            font=_font(11),
                            text_color=Theme.SUBTEXT,
                        ).grid(row=2, column=0, padx=10, pady=(0, 8), sticky="w")
                    elif segment.get("error"):
                        ctk.CTkLabel(
                            row,
                            text=str(segment.get("error")),
                            font=_font(11),
                            text_color=Theme.DANGER,
                        ).grid(row=2, column=0, padx=10, pady=(0, 8), sticky="w")

        summary_label = getattr(self, "route_schedule_summary_label", None)
        if summary_label is not None:
            summary = getattr(self, "current_route_summary", {}) or {}
            text = (
                f"Start: {self.current_route_start_time} · Fahrt: {summary.get('total_travel_minutes', 0)} min"
                f" · Aufenthalt: {summary.get('total_service_minutes', 0)} min"
                f" · Warten: {summary.get('total_wait_minutes', 0)} min"
            )
            if summary.get("end_time"):
                text += f" · Ende: {summary.get('end_time')}"
            if summary.get("has_conflicts"):
                text += " · Konflikte vorhanden"
            if status_text and not self.current_route_segments:
                text = status_text
            summary_label.configure(text=text)

    def _trigger_route_metrics_recalc(self, force_routing: bool = False):
        self._sync_current_route_stop_data_from_markers()
        nodes = self._build_route_nodes()
        if len(nodes) < 2:
            self.current_route_segments = []
            self.current_route_summary = {}
            self._update_route_metrics_ui("Noch keine Stopps geplant.")
            return

        if not force_routing:
            segments = []
            cache_missing = False
            for index in range(len(nodes) - 1):
                cache_key = f"{nodes[index].get('id', '')}->{nodes[index + 1].get('id', '')}"
                minutes = self.current_route_travel_time_cache.get(cache_key)
                distance = self.current_route_distance_cache.get(cache_key)
                if minutes is None:
                    cache_missing = True
                segments.append({
                    "cache_key": cache_key,
                    "travel_minutes": minutes,
                    "distance_km": distance,
                })
            self._apply_schedule_results(segments)
            self._update_route_metrics_ui("Berechne Fahrzeiten …" if cache_missing else "")
            if not cache_missing:
                return

        self._route_metrics_job += 1
        job_id = self._route_metrics_job
        nodes_snapshot = [dict(node) for node in nodes]
        cache_snapshot = dict(self.current_route_travel_time_cache)
        distance_snapshot = dict(self.current_route_distance_cache)
        route_mode = self.current_route_route_mode
        self._update_route_metrics_ui("Berechne Fahrzeiten …")

        def _worker():
            results = []
            updated_cache = dict(cache_snapshot)
            updated_distance = dict(distance_snapshot)
            for index in range(len(nodes_snapshot) - 1):
                segment = get_travel_segment(nodes_snapshot[index], nodes_snapshot[index + 1], updated_cache, route_mode=route_mode)
                cache_key = segment.get("cache_key") or f"{nodes_snapshot[index].get('id')}->{nodes_snapshot[index + 1].get('id')}"
                minutes = segment.get("minutes")
                if minutes is None:
                    minutes = updated_cache.get(cache_key)
                distance = segment.get("distance_km")
                if distance is None:
                    distance = updated_distance.get(cache_key)
                if distance is None:
                    distance = estimate_distance_km(nodes_snapshot[index], nodes_snapshot[index + 1])
                if minutes is not None:
                    updated_cache[cache_key] = minutes
                if distance is not None:
                    updated_distance[cache_key] = distance
                results.append({
                    "cache_key": cache_key,
                    "travel_minutes": minutes,
                    "distance_km": distance,
                    "error": segment.get("error", ""),
                })

            def _apply():
                if job_id != self._route_metrics_job:
                    return
                self.current_route_travel_time_cache = updated_cache
                self.current_route_distance_cache = updated_distance
                self._apply_schedule_results(results)
                status = ""
                if any(segment.get("travel_minutes") is None for segment in results):
                    status = "Einzelne Fahrzeiten konnten nicht geladen werden."
                self._update_route_metrics_ui(status)

            self.after(0, _apply)

        threading.Thread(target=_worker, daemon=True).start()

    # ---------- Marker helpers ----------
    def _style_marker_label(self, marker):
        marker.set_text(" ")
        try:
            if hasattr(marker, "label"):
                marker.label.configure(
                    fg_color=Theme.resolve(Theme.PANEL),
                    text_color=Theme.resolve(Theme.TEXT),
                )
        except Exception:
            pass
        marker.set_text("")

    def _alpha_label_for_index(self, index: int) -> str:
        number = int(index or 0)
        if number <= 0:
            return ""
        letters = []
        while number > 0:
            number, remainder = divmod(number - 1, 26)
            letters.append(chr(ord("A") + remainder))
        return "".join(reversed(letters))

    def _apply_route_marker_labels(self):
        labels_by_marker = {}
        if self.current_tour_id:
            regular_markers = [m for m in getattr(self, "route_markers", []) if m and not self._is_depot_marker(m)]
            for index, marker in enumerate(regular_markers, start=1):
                labels_by_marker[id(marker)] = self._alpha_label_for_index(index)

        for marker in getattr(self, "marker_list", []):
            if marker is None:
                continue
            try:
                marker.set_text(labels_by_marker.get(id(marker), ""))
            except Exception:
                pass

    def _format_products_block(self, products: str) -> str:
        product_items = [part.strip() for part in str(products or "").split(";") if str(part or "").strip()]
        if not product_items:
            return ""
        product_lines = []
        for item in product_items:
            text = str(item or "").strip()
            quantity = "1"
            match = PRODUCT_QTY_PREFIX_RE.match(text)
            if match:
                quantity = str(match.group(1) or "1")
                text = str(match.group(2) or "").strip() or text
            product_lines.append(f"{quantity}x {text}")
        return "\nProdukte:\n" + "\n".join(product_lines) + "\n\n"

    def _build_popup_text(self, d: dict) -> str:
        payload = _normalize_order_and_delivery_addresses(d)
        notes = (d.get("Notizen", "") or "").strip()
        products_line = self._format_products_block(d.get("Produkte", ""))
        notes_line = f"\n\nNotizen:\n{notes}" if notes else ""
        order_address = _format_order_address(payload) or "Nicht hinterlegt"
        delivery_address = _format_delivery_address_multiline(payload) or "Nicht hinterlegt"
        total_product_weight = str(d.get("Gewicht", "")).strip() or "N/A"
        order_weight = str(d.get("Auftragsgewicht", "")).strip()
        order_weight_line = ""
        if order_weight and order_weight != total_product_weight:
            order_weight_line = f"Auftragsgewicht: {order_weight}\n"
        return (
            f"Auftragsadresse: {order_address}\n"
            f"\n"
            f"Lieferadresse:\n{delivery_address}\n\n"
            f"Auftragsnummer: {d.get('Auftragsnummer', '')}\n"
            f"Bestelldatum: {d.get('Bestelldatum', '')}\n"
            f"Gesamtgewicht Produkte: {total_product_weight}\n"
            f"{order_weight_line}"
            f"{products_line}"
            f"Lieferart: {_normalize_delivery_type(d.get('Lieferart', DEFAULT_DELIVERY_TYPE))}\n"
            f"Status: {d.get('Status', 'nicht festgelegt')}\n\n"
            f"Email: {d.get('Email', '')}\n"
            f"Telefon: {d.get('Telefon', '')}"
            f"{notes_line}"
        )

    def _marker_key(self, marker):
        anchor = getattr(marker, "route_anchor_id", None)
        if anchor:
            return ("anchor", str(anchor))
        import_id = str(getattr(marker, "import_id", "")).strip()
        if import_id:
            return ("import", import_id)
        auftrag = str(getattr(marker, "auftragsnummer", "")).strip()
        if auftrag and auftrag.upper() != "N/A":
            return ("auftrag", auftrag)
        try:
            lat, lng = marker.position[0], marker.position[1]
            return ("coord", round(float(lat), 7), round(float(lng), 7))
        except Exception:
            return ("unknown",)

    def _item_identity_key(self, item: dict | None):
        data = item if isinstance(item, dict) else {}
        import_id = str(data.get("ImportID") or "").strip()
        if import_id:
            return ("import", import_id)
        order_number = str(data.get("Auftragsnummer") or "").strip()
        if order_number:
            return ("auftrag", order_number)
        return None

    def _collect_non_system_markers_by_identity(self) -> tuple[dict, list]:
        marker_by_identity = {}
        markers_without_identity = []
        for marker in list(self.marker_list):
            if getattr(marker, "is_system", False):
                continue
            marker_data = dict(getattr(marker, "data", {}) or {})
            identity_key = self._item_identity_key(marker_data)
            if identity_key is None:
                markers_without_identity.append(marker)
                continue
            marker_by_identity[identity_key] = marker
        return marker_by_identity, markers_without_identity

    def _apply_payload_to_existing_marker(self, marker, item: dict):
        item = _normalize_order_and_delivery_addresses(item)
        try:
            marker.set_position(float(item["lat"]), float(item["lng"]))
        except Exception:
            pass
        marker.data = dict(item)
        marker.status = str(item.get("Status") or "nicht festgelegt")
        marker.email = str(item.get("Email") or "")
        marker.auftragsnummer = str(item.get("Auftragsnummer") or "").strip()
        marker.import_id = str(item.get("ImportID") or "").strip()
        marker.full_info = self._build_popup_text(marker.data)

    def _delete_marker_safely(self, marker, *, remove_from_list: bool = True):
        try:
            marker.delete()
        except Exception:
            pass
        if not remove_from_list:
            return
        try:
            self.marker_list.remove(marker)
        except ValueError:
            pass

    def _warn_sql_cache_save_error(self, cache_error: str | None):
        if cache_error:
            messagebox.showwarning("Import", f"Geocode-Cache konnte nicht gespeichert werden:\n{cache_error}")

    def _refresh_after_sql_marker_changes(self, has_changes: bool, *, show_map: bool):
        if has_changes:
            self.save_pins()
            self._refresh_all_markers()
            if show_map:
                self.show_page("map")
        self._refresh_pages("list", "nonmap", visible_only=False)

    def _fail_sql_import(self, error: Exception | str, *, message: str):
        error_text = str(error)
        self._set_sql_import_status(f"Fehler: {error_text}", running=False)
        messagebox.showerror("Fehler", f"{message}:\n{error_text}")

    # ---------- Route Panel ----------
    def refresh_route_panel(self):
        tree = getattr(self, "route_tree", None)
        if tree is None:
            return

        for item in tree.get_children():
            tree.delete(item)

        total = 0.0
        stop_map = {self._marker_key(m): stop for m, stop in zip([m for m in self.route_markers if m and not self._is_depot_marker(m)], self.current_route_stop_data)}
        tree.tag_configure("conflict", background=Theme.resolve(Theme.CONFLICT_BG))
        tree.tag_configure("normal", background=Theme.resolve(Theme.PANEL))
        for idx, m in enumerate(getattr(self, "route_markers", []), start=1):
            d = getattr(m, "data", {}) or {}
            name = str(d.get("Name", "")).strip()
            street = str(d.get("Strasse", "")).strip()
            plz = str(d.get("PLZ", "")).strip()
            ort = str(d.get("Ort", "")).strip()
            line2_parts = [p for p in [street, " ".join([plz, ort]).strip()] if p]
            line2 = ", ".join(line2_parts).strip()

            if name and line2:
                address = f"{name} / {line2}"
            elif name:
                address = name
            elif line2:
                address = line2
            else:
                try:
                    address = f"{m.position[0]:.5f}, {m.position[1]:.5f}"
                except Exception:
                    address = "Unbekannt"

            if getattr(m, "is_system", False):
                w_text = ""
                time_window = "—"
                service_text = "—"
                eta = ""
                etd = ""
                tags = ("normal",)
            else:
                w_val = self._marker_weight_value(m)
                total += w_val
                w_text = self._format_weight(w_val)
                stop = stop_map.get(self._marker_key(m), {})
                time_window = self._format_time_window(stop)
                service_text = f"{int(stop.get('service_minutes') or 0)} min"
                eta = str(stop.get("planned_arrival") or "")
                etd = str(stop.get("planned_departure") or "")
                tags = ("conflict",) if stop.get("schedule_conflict") else ("normal",)

            key = self._marker_key(m)
            tree.insert("", "end", values=(idx, address, time_window, service_text, eta, etd, w_text), iid=str(key), tags=tags)

        lbl = getattr(self, "route_total_label", None)
        if lbl is not None:
            lbl.configure(text=f"Totalgewicht: {self._format_weight(total)}")

    def _apply_route_order_and_recalc(self):
        self._rebuild_route_from_markers()
        self._trigger_route_metrics_recalc(force_routing=True)

    def _get_selected_route_key_from_panel(self):
        tree = getattr(self, "route_tree", None)
        if tree is None:
            return None
        sel = tree.focus()
        return sel if sel else None

    def _get_route_marker_by_key(self, route_key):
        key_text = str(route_key) if route_key is not None else ""
        for marker in getattr(self, "route_markers", []):
            if str(self._marker_key(marker)) == key_text:
                return marker
        return None

    def _is_route_tree_key_draggable(self, route_key) -> bool:
        marker = self._get_route_marker_by_key(route_key)
        return bool(marker is not None and not self._is_depot_marker(marker))

    def _tree_row_insert_after(self, tree, route_key, y_pos: int) -> bool:
        try:
            bbox = tree.bbox(route_key)
        except tk.TclError:
            bbox = ()
        if not bbox:
            return True
        row_y = bbox[1]
        row_height = bbox[3]
        return y_pos >= (row_y + (row_height / 2))

    def _reorder_route_stop_by_keys(self, source_key, target_key, insert_after: bool) -> bool:
        source_marker = self._get_route_marker_by_key(source_key)
        if source_marker is None or self._is_depot_marker(source_marker):
            return False

        regular_markers = [marker for marker in self.route_markers if marker and not self._is_depot_marker(marker)]
        if len(regular_markers) < 2:
            return False

        source_key_text = str(source_key)
        remaining = [marker for marker in regular_markers if str(self._marker_key(marker)) != source_key_text]
        if len(remaining) == len(regular_markers):
            return False

        target_marker = self._get_route_marker_by_key(target_key)
        if target_marker is None:
            insert_index = len(remaining)
        elif self._is_depot_marker(target_marker):
            insert_index = len(remaining) if getattr(target_marker, "route_anchor_id", None) == "depot_end" else 0
        else:
            target_key_text = str(target_key)
            target_index = next(
                (index for index, marker in enumerate(remaining) if str(self._marker_key(marker)) == target_key_text),
                None,
            )
            if target_index is None:
                insert_index = len(remaining)
            else:
                insert_index = target_index + (1 if insert_after else 0)

        remaining.insert(max(0, min(insert_index, len(remaining))), source_marker)
        self.route_markers = remaining
        self._apply_route_order_and_recalc()
        return True

    def on_route_tree_drag_start(self, event):
        tree = getattr(self, "route_tree", None)
        if tree is None or tree.identify_region(event.x, event.y) != "cell":
            self._route_drag_key = None
            self._route_drag_target = None
            return

        route_key = tree.identify_row(event.y)
        if not route_key or not self._is_route_tree_key_draggable(route_key):
            self._route_drag_key = None
            self._route_drag_target = None
            return

        self._route_drag_key = route_key
        self._route_drag_target = route_key
        try:
            tree.selection_set(route_key)
            tree.focus(route_key)
        except tk.TclError:
            self._route_drag_key = None
            self._route_drag_target = None

    def on_route_tree_drag_motion(self, event):
        tree = getattr(self, "route_tree", None)
        if tree is None or not self._route_drag_key:
            return

        route_key = tree.identify_row(event.y)
        if route_key:
            self._route_drag_target = route_key

    def on_route_tree_drag_release(self, event):
        tree = getattr(self, "route_tree", None)
        source_key = self._route_drag_key
        if tree is None or not source_key:
            self._route_drag_key = None
            self._route_drag_target = None
            return

        target_key = tree.identify_row(event.y) or self._route_drag_target
        self._route_drag_key = None
        self._route_drag_target = None

        if not target_key or str(target_key) == str(source_key):
            return

        insert_after = self._tree_row_insert_after(tree, target_key, event.y)
        if not self._reorder_route_stop_by_keys(source_key, target_key, insert_after):
            return

        try:
            tree.selection_set(str(source_key))
            tree.focus(str(source_key))
        except tk.TclError:
            pass

    def on_route_tree_double_click(self, event):
        tree = getattr(self, "route_tree", None)
        if tree is None or tree.identify_region(event.x, event.y) != "cell":
            return

        route_key = tree.identify_row(event.y)
        if not route_key:
            return

        try:
            tree.selection_set(route_key)
            tree.focus(route_key)
        except tk.TclError:
            return

        marker = self._get_route_marker_by_key(route_key)
        if marker is None or self._is_depot_marker(marker):
            return

        self.open_selected_stop_editor()

    def move_selected_stop_up(self):
        sel = self._get_selected_route_key_from_panel()
        if not sel:
            return
        keys = [str(self._marker_key(m)) for m in self.route_markers]
        try:
            i = keys.index(sel)
        except ValueError:
            return
        if i <= 0:
            return
        self.route_markers[i - 1], self.route_markers[i] = self.route_markers[i], self.route_markers[i - 1]
        self._apply_route_order_and_recalc()
        try:
            self.route_tree.selection_set(sel)
            self.route_tree.focus(sel)
        except Exception:
            pass

    def move_selected_stop_down(self):
        sel = self._get_selected_route_key_from_panel()
        if not sel:
            return
        keys = [str(self._marker_key(m)) for m in self.route_markers]
        try:
            i = keys.index(sel)
        except ValueError:
            return
        if i >= len(self.route_markers) - 1:
            return
        self.route_markers[i + 1], self.route_markers[i] = self.route_markers[i], self.route_markers[i + 1]
        self._apply_route_order_and_recalc()
        try:
            self.route_tree.selection_set(sel)
            self.route_tree.focus(sel)
        except Exception:
            pass

    def remove_selected_stop_from_route_panel(self):
        sel = self._get_selected_route_key_from_panel()
        if not sel:
            return
        new_list = [m for m in self.route_markers if str(self._marker_key(m)) != sel]
        if len(new_list) == len(self.route_markers):
            return
        self.route_markers = new_list
        self._apply_route_order_and_recalc()

    def apply_route_start_time(self, value: str):
        parsed = parse_time(value)
        if parsed is None:
            messagebox.showwarning("Zeitplan", "Bitte eine gültige Startzeit im Format HH:MM eingeben.")
            return
        self.current_route_start_time = format_time(parsed)
        _set_text_input_value(getattr(self, "route_start_time_entry", None), self.current_route_start_time)
        self._trigger_route_metrics_recalc(force_routing=False)

    def open_selected_stop_editor(self):
        sel = self._get_selected_route_key_from_panel()
        if not sel:
            messagebox.showwarning("Stopp", "Bitte zuerst einen Stopp auswählen.")
            return

        marker = next((m for m in self.route_markers if str(self._marker_key(m)) == sel), None)
        if marker is None or self._is_depot_marker(marker):
            messagebox.showwarning("Stopp", "Zeitfenster können nur für reguläre Stopps bearbeitet werden.")
            return

        stop = next((s for s in self.current_route_stop_data if self._stop_key(s) == self._marker_key(marker)), None)
        if not stop:
            stop = self._make_default_stop_from_marker(marker, order=1)

        dlg = ctk.CTkToplevel(self)
        dlg.title("Stopp bearbeiten")
        dlg.geometry("540x440")
        dlg.resizable(True, True)
        dlg.configure(fg_color=Theme.BG)
        dlg.attributes("-topmost", True)

        shell = ctk.CTkFrame(dlg, corner_radius=18, fg_color=Theme.PANEL, border_width=1, border_color=Theme.BORDER)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        shell.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(shell, text=stop.get("name") or "Stopp", font=_font(15, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=(16, 4), sticky="w"
        )
        ctk.CTkLabel(shell, text=stop.get("address") or "", font=_font(12), text_color=Theme.SUBTEXT).grid(
            row=1, column=0, padx=16, pady=(0, 10), sticky="w"
        )

        fields = ctk.CTkFrame(shell, fg_color="transparent")
        fields.grid(row=2, column=0, padx=16, pady=(0, 10), sticky="ew")
        fields.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(fields, text="Zeitfenster Start", font=_font(12, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=(0, 8), pady=(0, 6), sticky="w"
        )
        ctk.CTkLabel(fields, text="Zeitfenster Ende", font=_font(12, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=1, padx=(8, 0), pady=(0, 6), sticky="w"
        )

        start_entry = TimeInput(fields, height=36)
        start_entry.grid(row=1, column=0, padx=(0, 8), pady=4, sticky="ew")
        start_entry.set(stop.get("time_window_start"))

        end_entry = TimeInput(fields, height=36)
        end_entry.grid(row=1, column=1, padx=(8, 0), pady=4, sticky="ew")
        end_entry.set(stop.get("time_window_end"))

        ctk.CTkLabel(shell, text="Aufenthaltszeit", font=_font(12, "bold"), text_color=Theme.TEXT).grid(
            row=3, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        service_menu = ctk.CTkOptionMenu(
            shell,
            values=SERVICE_MINUTE_OPTIONS,
            height=36,
            corner_radius=12,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL_2,
            button_color=Theme.ACCENT,
            button_hover_color=Theme.ACCENT_HOVER,
            text_color=Theme.TEXT,
        )
        service_menu.grid(row=4, column=0, padx=16, pady=(0, 8), sticky="ew")
        service_menu.set(str(int(stop.get("service_minutes") or 0)))

        info_label = ctk.CTkLabel(shell, text="", font=_font(12), text_color=Theme.SUBTEXT, justify="left")
        info_label.grid(row=5, column=0, padx=16, pady=(0, 8), sticky="w")

        debounce_ref = {"job": None}

        def _apply_preview():
            valid, error = validate_time_window(start_entry.get().strip(), end_entry.get().strip())
            if not valid:
                info_label.configure(text=error, text_color=Theme.DANGER)
                return

            stop["time_window_start"] = start_entry.get().strip()
            stop["time_window_end"] = end_entry.get().strip()
            stop["service_minutes"] = int(service_menu.get() or 0)
            info_label.configure(text="Zeitfenster gültig. Zeitplan wird aktualisiert.", text_color=Theme.SUBTEXT)
            self._sync_current_route_stop_data_from_markers(preferred_stops=[stop])
            self._trigger_route_metrics_recalc(force_routing=False)

        def _debounced_preview(event=None):
            if debounce_ref["job"] is not None:
                try:
                    dlg.after_cancel(debounce_ref["job"])
                except Exception:
                    pass
            debounce_ref["job"] = dlg.after(350, _apply_preview)

        start_entry.bind_change(_debounced_preview)
        end_entry.bind_change(_debounced_preview)
        service_menu.configure(command=lambda _value: _debounced_preview())

        btns = ctk.CTkFrame(shell, fg_color="transparent")
        btns.grid(row=6, column=0, padx=16, pady=(6, 16), sticky="ew")
        btns.grid_columnconfigure((0, 1), weight=1)

        def _save():
            valid, error = validate_time_window(start_entry.get().strip(), end_entry.get().strip())
            if not valid:
                messagebox.showwarning("Stopp", error)
                return
            stop["time_window_start"] = start_entry.get().strip()
            stop["time_window_end"] = end_entry.get().strip()
            stop["service_minutes"] = int(service_menu.get() or 0)
            self._sync_current_route_stop_data_from_markers(preferred_stops=[stop])
            self._trigger_route_metrics_recalc(force_routing=False)
            try:
                dlg.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            btns,
            text="Schließen",
            height=40,
            corner_radius=14,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=dlg.destroy,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            btns,
            text="Speichern",
            height=40,
            corner_radius=14,
            font=_font(13, "bold"),
            fg_color=Theme.SUCCESS,
            hover_color=Theme.SUCCESS_HOVER,
            command=_save,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        _apply_preview()
        dlg.grab_set()
        dlg.focus_force()

    def focus_stop_from_route_panel(self):
        sel = self._get_selected_route_key_from_panel()
        if not sel:
            return
        for m in self.route_markers:
            if str(self._marker_key(m)) == sel:
                self.on_marker_click(m)
                try:
                    lat, lng = m.position[0], m.position[1]
                    self.map_widget.set_position(lat, lng)
                    z = self._get_current_zoom()
                    if z is not None and z < 14:
                        self.map_widget.set_zoom(14)
                except Exception:
                    pass
                return

    # ---------- Tours helper ----------
    def _build_tour_usage_cache(self) -> dict:
        if self._tour_usage_cache_revision == self._tour_data_revision and isinstance(self._tour_usage_cache, dict):
            return self._tour_usage_cache

        usage = {}
        for tour in self._load_tours():
            for stop in tour.get("stops", []):
                if not isinstance(stop, dict):
                    continue
                stop_key = None
                auftrag = str(stop.get("auftragsnummer", "")).strip()
                if auftrag and auftrag.upper() != "N/A":
                    stop_key = ("auftrag", auftrag)
                else:
                    lat = stop.get("lat")
                    lng = stop.get("lng")
                    if lat is not None and lng is not None:
                        try:
                            stop_key = ("coord", round(float(lat), 7), round(float(lng), 7))
                        except Exception:
                            stop_key = None
                if stop_key is not None and stop_key not in usage:
                    usage[stop_key] = tour

        self._tour_usage_cache = usage
        self._tour_usage_cache_revision = self._tour_data_revision
        return usage

    def _pin_used_in_any_tour_by_key(self, key):
        if not isinstance(key, tuple) or not key:
            return None
        usage = self._build_tour_usage_cache()
        return usage.get(key)

    def show_toast(self, text: str, duration_ms: int = 2200):
        try:
            if getattr(self, "_toast_win", None) is not None:
                try:
                    self._toast_win.destroy()
                except Exception:
                    pass
                self._toast_win = None

            win = ctk.CTkToplevel(self)
            self._toast_win = win
            win.overrideredirect(True)
            win.attributes("-topmost", True)
            try:
                win.attributes("-alpha", 0.96)
            except Exception:
                pass

            card = ctk.CTkFrame(
                win,
                corner_radius=18,
                fg_color=Theme.OVERLAY_PANEL,
                border_width=1,
                border_color=Theme.OVERLAY_BORDER,
            )
            card.pack(fill="both", expand=True)

            lbl = ctk.CTkLabel(card, text=text, font=_font(13, "bold"), text_color=Theme.TEXT)
            lbl.pack(padx=16, pady=12)

            self.update_idletasks()
            win.update_idletasks()

            w = 260
            h = 52
            x_target = self.winfo_rootx() + self.winfo_width() - w - 24
            y_target = self.winfo_rooty() + self.winfo_height() - h - 24
            y_start = y_target + 18

            win.geometry(f"{w}x{h}+{x_target}+{y_start}")

            steps = 10
            step_ms = 12
            dy = (y_start - y_target) / steps

            def _slide(i=0):
                if not win.winfo_exists():
                    return
                y = int(y_start - dy * i)
                win.geometry(f"{w}x{h}+{x_target}+{y}")
                if i < steps:
                    win.after(step_ms, lambda: _slide(i + 1))

            _slide()

            def _close():
                if getattr(self, "_toast_win", None) is not None:
                    try:
                        self._toast_win.destroy()
                    except Exception:
                        pass
                    self._toast_win = None

            win.after(duration_ms, _close)
        except Exception:
            pass

    def _pin_used_in_any_tour(self, marker):
        return self._pin_used_in_any_tour_by_key(self._marker_key(marker))

    # ---------- Marker Icons ----------
    def _make_circle_icon(self, hex_color: str, size: int) -> ImageTk.PhotoImage:
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        pad = 1
        draw.ellipse((pad, pad, size - pad - 1, size - pad - 1), fill=hex_color, outline="#222222")
        return ImageTk.PhotoImage(img)

    def _make_square_icon(self, hex_color: str, size: int) -> ImageTk.PhotoImage:
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        pad = 2
        draw.rectangle((pad, pad, size - pad - 1, size - pad - 1), fill=hex_color, outline="#222222")
        return ImageTk.PhotoImage(img)

    def _make_triangle_icon(self, hex_color: str, size: int) -> ImageTk.PhotoImage:
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        pad = 1
        points = [
            (size // 2, pad),
            (size - pad - 1, size - pad - 1),
            (pad, size - pad - 1),
        ]
        draw.polygon(points, fill=hex_color, outline="#222222")
        return ImageTk.PhotoImage(img)

    def _delivery_shape(self, delivery_type: str) -> str:
        normalized = _normalize_delivery_type(delivery_type)
        if normalized == "Mit Verteilung":
            return "square"
        if normalized == "mit Verteilung & montage":
            return "triangle"
        return "circle"

    def _make_delivery_icon(self, delivery_type: str, hex_color: str, size: int) -> ImageTk.PhotoImage:
        shape = self._delivery_shape(delivery_type)
        if shape == "square":
            return self._make_square_icon(hex_color, size)
        if shape == "triangle":
            return self._make_triangle_icon(hex_color, size)
        return self._make_circle_icon(hex_color, size)

    def _build_marker_icons(self, size: int) -> dict:
        icons = {}
        for delivery_type in DELIVERY_TYPE_OPTIONS:
            normalized = _normalize_delivery_type(delivery_type)
            for status, color in self.status_colors.items():
                icons[(normalized, status)] = self._make_delivery_icon(normalized, color, size)
            icons[(normalized, "tour")] = self._make_delivery_icon(normalized, self.tour_pin_color, size)
        return icons

    def _get_marker_icon(self, status: str, in_tour: bool, delivery_type: str = DEFAULT_DELIVERY_TYPE) -> ImageTk.PhotoImage:
        normalized_delivery_type = _normalize_delivery_type(delivery_type)
        status_key = "tour" if in_tour else str(status or "").strip()
        icon = self.marker_icons.get((normalized_delivery_type, status_key))
        if icon is not None:
            return icon
        fallback = self.marker_icons.get((normalized_delivery_type, "nicht festgelegt"))
        if fallback is not None:
            return fallback
        # Ultimate fallback should always exist for default delivery type.
        return self.marker_icons[(DEFAULT_DELIVERY_TYPE, "nicht festgelegt")]

    def _get_depot_marker_icon(self, anchor_id: str) -> ImageTk.PhotoImage:
        cache_key = (str(anchor_id), int(self.marker_icon_size))
        cached = self._depot_icon_cache.get(cache_key)
        if cached is not None:
            return cached
        color = "#10B981" if str(anchor_id) == "depot_start" else "#EF4444"
        icon = self._make_circle_icon(color, self.marker_icon_size)
        self._depot_icon_cache[cache_key] = icon
        return icon

    # ---------- Depot / Systemmarker ----------
    def _is_depot_marker(self, marker) -> bool:
        if marker is None:
            return False
        return bool(getattr(marker, "is_system", False) and getattr(marker, "route_anchor_id", None) in ("depot_start", "depot_end"))

    def _ensure_base_latlng(self) -> bool:
        if self.base_latlng is not None:
            return True
        try:
            latlng = self.geocoding_service.lookup(self.base_address)
            if not latlng:
                messagebox.showwarning("Depot", f"Depot-Adresse nicht gefunden:\n{self.base_address}")
                return False
            self.base_latlng = latlng
            self.geocoding_service.save_cache()
            return True
        except OSError as e:
            logger.exception("Depot cache could not be saved.")
            messagebox.showwarning("Depot", f"Depot konnte nicht geladen werden:\n{e}")
            return False
        except Exception as e:
            logger.exception("Depot geocoding failed.")
            messagebox.showwarning("Depot", f"Depot konnte nicht geladen werden:\n{e}")
            return False

    def _get_or_create_depot_anchor(self, anchor_id: str, label: str):
        for m in getattr(self, "marker_list", []):
            if getattr(m, "route_anchor_id", None) == anchor_id:
                return m

        if not self._ensure_base_latlng():
            return None

        lat, lng = self.base_latlng
        data = {
            "Name": label,
            "Strasse": "Konstanzerstrasse 14",
            "PLZ": "8274",
            "Ort": "Tägerwilen",
            "Auftragsnummer": f"DEPOT-{anchor_id}",
            "Bestelldatum": "",
            "Gewicht": "",
            "Email": "",
            "Telefon": "",
            "Status": "nicht festgelegt",
        }

        depot_color = "#10B981" if anchor_id == "depot_start" else "#EF4444"
        icon = self._make_circle_icon(depot_color, self.marker_icon_size)

        m = self.map_widget.set_marker(lat, lng, text="", icon=icon, command=self.on_marker_click)
        m.data = data
        m.status = "nicht festgelegt"
        m.full_info = self._build_popup_text(data)
        m.email = ""
        m.auftragsnummer = data["Auftragsnummer"]
        m.is_system = True
        m.route_anchor_id = anchor_id

        self._style_marker_label(m)
        self.marker_list.append(m)

        if anchor_id == "depot_start":
            self.depot_start_marker = m
        elif anchor_id == "depot_end":
            self.depot_end_marker = m

        return m

    def _ensure_depot_markers_exist(self):
        if self.depot_start_marker and self.depot_end_marker:
            return
        if not self._ensure_base_latlng():
            return
        if not self.depot_start_marker:
            self.depot_start_marker = self._get_or_create_depot_anchor("depot_start", "Start (Depot)")
        if not self.depot_end_marker:
            self.depot_end_marker = self._get_or_create_depot_anchor("depot_end", "Ende (Depot)")

    # ---------- Marker creation ----------
    def _create_marker(self, lat, lng, data: dict, status: str):
        if isinstance(data, dict):
            data = _normalize_order_and_delivery_addresses(data)
            data["Lieferart"] = _normalize_delivery_type(data.get("Lieferart"))
            if _is_non_map_delivery_type(data.get("Lieferart")):
                return None

        dummy_marker = type("Dummy", (), {})()
        dummy_marker.data = data
        dummy_marker.auftragsnummer = data.get("Auftragsnummer", "")
        dummy_marker.position = (lat, lng)
        in_tour = self._pin_used_in_any_tour(dummy_marker) is not None
        delivery_type = _normalize_delivery_type(data.get("Lieferart"))

        icon = self._get_marker_icon(status, in_tour, delivery_type=delivery_type)
        m = self.map_widget.set_marker(lat, lng, text="", icon=icon, command=self.on_marker_click)

        m.data = data
        m.status = status
        m.full_info = self._build_popup_text(data)
        m.email = data.get("Email", "")
        m.auftragsnummer = data.get("Auftragsnummer", "")
        m.import_id = str(data.get("ImportID") or "").strip()

        self._style_marker_label(m)
        self.marker_list.append(m)
        return m

    # ---------- Route ----------
    def add_current_to_route(self):
        if not self.current_selected_marker:
            return
        m = self.current_selected_marker
        if self._is_depot_marker(m):
            return
        if m in self.route_markers:
            messagebox.showinfo("Route", "Dieser Stopp ist bereits in der aktuellen Route.")
            return

        used_in = self._pin_used_in_any_tour(m)
        if used_in is not None:
            tour_name = (used_in.get("name") or "").strip()
            tour_date = _normalize_date_string(used_in.get("date"))
            label = f"{tour_date}" + (f" – {tour_name}" if tour_name else "")
            messagebox.showwarning("Tour-Planung", f"Dieser Pin ist bereits in einer anderen Liefertour eingeplant:\n{label}")
            return

        self._ensure_depot_markers_exist()

        if not self.route_markers:
            self.route_markers = [self.depot_start_marker, self.depot_end_marker]

        if self.depot_end_marker in self.route_markers:
            idx_end = self.route_markers.index(self.depot_end_marker)
            self.route_markers.insert(idx_end, m)
        else:
            self.route_markers.append(m)

        self._rebuild_route_from_markers()
        self._trigger_route_metrics_recalc(force_routing=True)

    def optimize_current_route_order(self):
        self._ensure_depot_markers_exist()
        regular_markers = [m for m in self.route_markers if m and not self._is_depot_marker(m)]
        if len(regular_markers) < 2:
            messagebox.showinfo("Route-Optimierung", "Für die Optimierung werden mindestens 2 Stopps benötigt.")
            return

        start_node = {
            "id": "depot_start",
            "name": "Start (Depot)",
            "lat": self.depot_start_marker.position[0] if self.depot_start_marker else None,
            "lon": self.depot_start_marker.position[1] if self.depot_start_marker else None,
        }
        end_node = {
            "id": "depot_end",
            "name": "Ende (Depot)",
            "lat": self.depot_end_marker.position[0] if self.depot_end_marker else None,
            "lon": self.depot_end_marker.position[1] if self.depot_end_marker else None,
        }

        stop_payload = []
        marker_by_uid = {}
        existing_map = {
            self._stop_key(stop): dict(stop)
            for stop in getattr(self, "current_route_stop_data", [])
            if isinstance(stop, dict)
        }
        for index, marker in enumerate(regular_markers, start=1):
            marker_key = self._marker_key(marker)
            base = existing_map.get(marker_key)
            stop = self._make_default_stop_from_marker(marker, existing=base, order=index)
            uid = f"marker:{index}:{marker_key}"
            stop["_optimizer_uid"] = uid
            marker_by_uid[uid] = marker
            stop_payload.append(stop)

        optimization = optimize_stop_order(
            start_node,
            stop_payload,
            end_node,
            start_time=self.current_route_start_time,
            travel_time_cache=self.current_route_travel_time_cache,
        )
        optimized_stops = optimization.get("stops") or []
        ordered_markers = []
        preferred_stops = []
        seen = set()

        for order, stop in enumerate(optimized_stops, start=1):
            uid = str(stop.get("_optimizer_uid") or "")
            marker = marker_by_uid.get(uid)
            if marker is None:
                continue
            marker_id = id(marker)
            if marker_id in seen:
                continue
            seen.add(marker_id)
            ordered_markers.append(marker)
            cleaned = dict(stop)
            cleaned.pop("_optimizer_uid", None)
            cleaned["order"] = order
            preferred_stops.append(cleaned)

        if len(ordered_markers) != len(regular_markers):
            fallback = [m for m in regular_markers if id(m) not in seen]
            ordered_markers.extend(fallback)
            for marker in fallback:
                preferred_stops.append(self._make_default_stop_from_marker(marker, order=len(preferred_stops) + 1))

        if self.depot_start_marker and self.depot_end_marker:
            self.route_markers = [self.depot_start_marker] + ordered_markers + [self.depot_end_marker]
        else:
            self.route_markers = ordered_markers

        self._sync_current_route_stop_data_from_markers(preferred_stops=preferred_stops)
        self._rebuild_route_from_markers()
        self._trigger_route_metrics_recalc(force_routing=True)

        best_metrics = optimization.get("metrics") or {}
        baseline_metrics = optimization.get("baseline_metrics") or {}
        before_drive = int(baseline_metrics.get("drive_minutes") or 0)
        after_drive = int(best_metrics.get("drive_minutes") or 0)
        before_late = int(baseline_metrics.get("late_minutes") or 0)
        after_late = int(best_metrics.get("late_minutes") or 0)
        messagebox.showinfo(
            "Route-Optimierung",
            (
                "Stoppreihenfolge optimiert.\n"
                f"Fahrzeit (geschätzt): {before_drive} min -> {after_drive} min\n"
                f"Zeitfenster-Verstöße (geschätzt): {before_late} min -> {after_late} min"
            ),
        )

    def calculate_route(self):
        if len(self.route_points) < 2:
            return
        self._route_request_job += 1
        job_id = self._route_request_job
        route_points = [tuple(point) for point in self.route_points]

        def _worker():
            try:
                route_list = fetch_route_path(route_points)
            except RouteServiceError as exc:
                message = str(exc)
                logger.warning("Route path could not be loaded: %s", exc)
                self.after(0, lambda: self._on_route_path_error(job_id, message))
                return
            except Exception:
                logger.exception("Unexpected route error")
                self.after(0, lambda: self._on_route_path_error(job_id, "Unerwarteter Fehler beim Routing."))
                return

            self.after(0, lambda: self._apply_route_path(job_id, route_list))

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_route_path(self, job_id: int, route_list: list):
        if job_id != self._route_request_job:
            return
        if self.route_path:
            try:
                self.route_path.delete()
            except AttributeError:
                logger.warning("Existing route path no longer supports delete().")
            except tk.TclError:
                logger.warning("Existing route path could not be removed from the map.")
        self.route_path = self.map_widget.set_path(route_list, color="#3498db", width=5)

    def _on_route_path_error(self, job_id: int, message: str):
        if job_id != self._route_request_job:
            return
        messagebox.showerror("Fehler", f"Route konnte nicht geladen werden: {message}")

    def clear_route(self):
        if self.route_path:
            try:
                self.route_path.delete()
            except Exception:
                pass
        self.route_path = None
        self.route_points = []
        self.route_markers = []
        self.current_route_stop_data = []
        self.current_route_segments = []
        self.current_route_travel_time_cache = {}
        self.current_route_distance_cache = {}
        messagebox.showinfo("Route", "Route wurde zurückgesetzt.")
        self._update_route_metrics_ui("Route wurde zurückgesetzt.")
        if hasattr(self, "lbl_current_tour"):
            self.lbl_current_tour.configure(text="(manuell)")

    def _rebuild_route_from_markers(self):
        if self.depot_start_marker and self.depot_end_marker:
            stops = [m for m in self.route_markers if m and not self._is_depot_marker(m)]
            self.route_markers = [self.depot_start_marker] + stops + [self.depot_end_marker]
        self._sync_current_route_stop_data_from_markers()
        self._apply_route_marker_labels()
        self.route_points = [m.position for m in self.route_markers if hasattr(m, "position")]

        if len(self.route_points) >= 2:
            self.calculate_route()
        else:
            if getattr(self, "route_path", None):
                try:
                    self.route_path.delete()
                except Exception:
                    pass
            self.route_path = None

    def export_route(self):
        if not self.route_points or len(self.route_points) < 2:
            messagebox.showwarning("Export", "Für den Export brauchst du mindestens 2 Stopps (Start + Ziel).")
            return

        origin = f"{self.route_points[0][0]},{self.route_points[0][1]}"
        destination = f"{self.route_points[-1][0]},{self.route_points[-1][1]}"
        intermediate = self.route_points[1:-1]

        max_waypoints = 9
        if len(intermediate) > max_waypoints:
            messagebox.showwarning("Export", f"Zu viele Zwischenstopps.\nIch exportiere nur die ersten {max_waypoints}.")
            intermediate = intermediate[:max_waypoints]

        waypoints_param = ""
        if intermediate:
            waypoints = "|".join([f"{p[0]},{p[1]}" for p in intermediate])
            waypoints_param = f"&waypoints={waypoints}"

        url = (
            "https://www.google.com/maps/dir/?api=1"
            f"&origin={origin}"
            f"&destination={destination}"
            f"{waypoints_param}"
            "&travelmode=driving"
        )
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Export", f"Konnte Google Maps nicht öffnen:\n{e}")

    # ---------- Marker click / status ----------
    def _set_info_card_text(self, text: str):
        widget = getattr(self, "info_label", None)
        value = str(text or "")
        if widget is None:
            return
        if isinstance(widget, tk.Text):
            try:
                widget.configure(state="normal")
                widget.delete("1.0", "end")
                for line in value.splitlines(keepends=True):
                    stripped = line.strip()
                    if stripped == "Produkte:" or stripped == "Lieferadresse:":
                        widget.insert("end", line, ("bold",))
                    else:
                        widget.insert("end", line)
                widget.configure(state="disabled")
                return
            except Exception:
                pass
        try:
            widget.configure(text=value)
        except Exception:
            pass

    def _set_tour_action_buttons_visible(self, visible: bool):
        show_btn = getattr(self, "btn_show_tour", None)
        remove_btn = getattr(self, "btn_remove_from_tour", None)
        if visible:
            if show_btn is not None:
                show_btn.grid()
            if remove_btn is not None:
                remove_btn.grid()
            return
        if show_btn is not None:
            show_btn.grid_remove()
        if remove_btn is not None:
            remove_btn.grid_remove()

    def _set_edit_order_button_visible(self, visible: bool):
        button = getattr(self, "btn_edit_order", None)
        if button is None:
            return
        if visible:
            button.grid()
            return
        button.grid_remove()

    def _show_info_card(self):
        try:
            self.info_card.grid(row=0, column=1, padx=(0, 14), pady=14, sticky="nsew")
        except Exception:
            pass

    def on_marker_click(self, marker):
        self.current_selected_marker = marker

        if hasattr(marker, "full_info"):
            self._set_info_card_text(marker.full_info)

        self._show_info_card()

        tour = self._pin_used_in_any_tour(marker)

        if hasattr(self, "status_menu"):
            if tour is not None and not self._is_depot_marker(marker):
                self.status_menu.set("Bereits eingeplant")
                self.status_menu.configure(state="disabled")
                self._selected_pin_tour = tour
                self._set_tour_action_buttons_visible(True)
            else:
                self.status_menu.configure(state="normal")
                self.status_menu.set(getattr(marker, "status", "nicht festgelegt"))
                self._selected_pin_tour = None
                self._set_tour_action_buttons_visible(False)

        # "Auftrag bearbeiten" nur bei normalen Pins
        self._set_edit_order_button_visible(bool(marker and not getattr(marker, "is_system", False)))

    def hide_info_card(self):
        try:
            self.info_card.grid_forget()
        except Exception:
            pass
        self._set_edit_order_button_visible(False)

    def set_selected_pin_status(self, status: str):
        marker = getattr(self, "current_selected_marker", None)
        if not marker:
            return

        if not self._is_depot_marker(marker):
            if self._pin_used_in_any_tour(marker) is not None:
                messagebox.showwarning("Status", "Dieser Pin ist bereits eingeplant und kann nicht geändert werden.")
                if hasattr(self, "status_menu"):
                    self.status_menu.set("Bereits eingeplant")
                    self.status_menu.configure(state="disabled")
                return

        if status not in self.status_colors:
            status = "nicht festgelegt"

        lat, lng = marker.position
        data = getattr(marker, "data", {}) or {}
        data["Status"] = status

        is_system = bool(getattr(marker, "is_system", False))
        anchor_id = getattr(marker, "route_anchor_id", None)

        try:
            marker.delete()
        except Exception:
            pass

        if marker in self.marker_list:
            self.marker_list.remove(marker)

        if is_system and anchor_id in ("depot_start", "depot_end"):
            new_marker = self._get_or_create_depot_anchor(anchor_id, data.get("Name", "Depot"))
            if new_marker:
                new_marker.data = data
                new_marker.status = status
                new_marker.full_info = self._build_popup_text(data)
        else:
            new_marker = self._create_marker(lat, lng, data, status)

        self.current_selected_marker = new_marker
        if new_marker:
            self._set_info_card_text(new_marker.full_info)

        self.save_pins()
        self._refresh_all_markers()

    def _map_filter_button_text(self) -> str:
        if bool(getattr(self, "map_filter_only_current_tour", False)):
            return "Filter: Nur aktuelle Tour"

        status_options = list(getattr(self, "map_filter_status_options", []))
        delivery_options = list(getattr(self, "map_filter_delivery_type_options", []))
        selected_statuses = {status for status in status_options if status in getattr(self, "map_filter_selected_statuses", set())}
        selected_delivery = {delivery for delivery in delivery_options if delivery in getattr(self, "map_filter_selected_delivery_types", set())}
        if len(selected_statuses) == len(status_options) and len(selected_delivery) == len(delivery_options):
            return "Filter: Alle"
        return f"Filter: {len(selected_statuses)} Status, {len(selected_delivery)} Lieferarten"

    def _update_map_filter_button(self):
        button = getattr(self, "map_filter_button", None)
        if button is None:
            return
        try:
            button.configure(text=self._map_filter_button_text())
        except Exception:
            pass

    def open_map_filter_dialog(self):
        existing = getattr(self, "_map_filter_dialog", None)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.focus_force()
                    return
            except Exception:
                pass

        dlg = ctk.CTkToplevel(self)
        dlg.title("Kartenfilter")
        dlg.geometry("420x560")
        dlg.resizable(False, False)
        dlg.configure(fg_color=Theme.BG)
        dlg.attributes("-topmost", True)
        self._map_filter_dialog = dlg

        shell = ctk.CTkFrame(dlg, corner_radius=16, fg_color=Theme.PANEL, border_width=1, border_color=Theme.BORDER)
        shell.pack(fill="both", expand=True, padx=12, pady=12)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(2, weight=1)
        shell.grid_rowconfigure(4, weight=1)

        ctk.CTkLabel(
            shell,
            text="Pins nach Status/Lieferart filtern",
            font=_font(14, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=0, column=0, padx=12, pady=(12, 8), sticky="w")

        only_current_tour_var = tk.BooleanVar(value=bool(self.map_filter_only_current_tour))
        ctk.CTkCheckBox(
            shell,
            text="Nur aktuelle Tour anzeigen",
            variable=only_current_tour_var,
            onvalue=True,
            offvalue=False,
            text_color=Theme.TEXT,
        ).grid(row=1, column=0, padx=12, pady=(0, 8), sticky="w")

        selected_statuses = set(getattr(self, "map_filter_selected_statuses", set()))
        status_vars = {}
        selected_delivery_types = {
            _normalize_delivery_type(value)
            for value in getattr(self, "map_filter_selected_delivery_types", set())
        }
        delivery_vars = {}
        status_list = ctk.CTkScrollableFrame(
            shell,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            **_scrollable_frame_kwargs(),
        )
        status_list.grid(row=2, column=0, padx=10, pady=(0, 8), sticky="nsew")
        status_list.grid_columnconfigure(0, weight=1)

        for row, status in enumerate(self.map_filter_status_options):
            var = tk.BooleanVar(value=status in selected_statuses)
            status_vars[status] = var
            ctk.CTkCheckBox(
                status_list,
                text=status,
                variable=var,
                onvalue=True,
                offvalue=False,
                text_color=Theme.TEXT,
            ).grid(row=row, column=0, padx=8, pady=6, sticky="w")

        ctk.CTkLabel(
            shell,
            text="Nach Lieferart filtern",
            font=_font(13, "bold"),
            text_color=Theme.TEXT,
        ).grid(row=3, column=0, padx=12, pady=(4, 6), sticky="w")

        delivery_list = ctk.CTkScrollableFrame(
            shell,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            **_scrollable_frame_kwargs(),
        )
        delivery_list.grid(row=4, column=0, padx=10, pady=(0, 8), sticky="nsew")
        delivery_list.grid_columnconfigure(0, weight=1)

        for row, delivery_type in enumerate(self.map_filter_delivery_type_options):
            normalized = _normalize_delivery_type(delivery_type)
            var = tk.BooleanVar(value=normalized in selected_delivery_types)
            delivery_vars[normalized] = var
            ctk.CTkCheckBox(
                delivery_list,
                text=normalized,
                variable=var,
                onvalue=True,
                offvalue=False,
                text_color=Theme.TEXT,
            ).grid(row=row, column=0, padx=8, pady=6, sticky="w")

        button_row = ctk.CTkFrame(shell, fg_color="transparent")
        button_row.grid(row=5, column=0, padx=10, pady=(0, 10), sticky="ew")
        button_row.grid_columnconfigure((0, 1, 2, 3), weight=1)

        def _set_all_statuses(value: bool):
            for var in status_vars.values():
                var.set(value)

        def _set_all_delivery_types(value: bool):
            for var in delivery_vars.values():
                var.set(value)

        def _close():
            popup = getattr(self, "_map_filter_dialog", None)
            self._map_filter_dialog = None
            if popup is None:
                return
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except Exception:
                pass

        def _apply():
            selected = {status for status, var in status_vars.items() if bool(var.get())}
            selected_delivery = {delivery for delivery, var in delivery_vars.items() if bool(var.get())}
            if not selected and not bool(only_current_tour_var.get()):
                selected = set(self.map_filter_status_options)
            if not selected_delivery and not bool(only_current_tour_var.get()):
                selected_delivery = {_normalize_delivery_type(value) for value in self.map_filter_delivery_type_options}
            self.map_filter_selected_statuses = selected
            self.map_filter_selected_delivery_types = selected_delivery
            self.map_filter_only_current_tour = bool(only_current_tour_var.get())
            self._update_map_filter_button()
            self._refresh_all_markers()
            _close()

        ctk.CTkButton(
            button_row,
            text="Alle",
            height=34,
            corner_radius=10,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: _set_all_statuses(True),
        ).grid(row=0, column=0, padx=(0, 4), sticky="ew")

        ctk.CTkButton(
            button_row,
            text="Keine Status",
            height=34,
            corner_radius=10,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: _set_all_statuses(False),
        ).grid(row=0, column=1, padx=4, sticky="ew")

        ctk.CTkButton(
            button_row,
            text="Alle Lieferarten",
            height=34,
            corner_radius=10,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: _set_all_delivery_types(True),
        ).grid(row=1, column=0, padx=(0, 4), pady=(6, 0), sticky="ew")

        ctk.CTkButton(
            button_row,
            text="Keine Lieferarten",
            height=34,
            corner_radius=10,
            fg_color=Theme.PANEL,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: _set_all_delivery_types(False),
        ).grid(row=1, column=1, padx=4, pady=(6, 0), sticky="ew")

        ctk.CTkButton(
            button_row,
            text="Übernehmen",
            height=34,
            corner_radius=10,
            fg_color=Theme.ACCENT,
            hover_color=Theme.ACCENT_HOVER,
            text_color=("white", "white"),
            command=_apply,
        ).grid(row=0, column=2, rowspan=2, padx=(4, 0), pady=(0, 0), sticky="nsew")

        dlg.protocol("WM_DELETE_WINDOW", _close)
        dlg.after(10, lambda: dlg.focus_force() if dlg.winfo_exists() else None)

    def delete_selected_pin(self):
        marker = getattr(self, "current_selected_marker", None)
        if not marker:
            messagebox.showwarning("Pin löschen", "Kein Pin ausgewählt.")
            return
        if getattr(marker, "is_system", False):
            messagebox.showwarning("Pin löschen", "System-Pins (Depot) können nicht gelöscht werden.")
            return
        if not messagebox.askyesno("Pin löschen", "Diesen Pin wirklich löschen?"):
            return
        try:
            marker.delete()
        except Exception as e:
            messagebox.showerror("Pin löschen", f"Pin konnte nicht gelöscht werden:\n{e}")
            return

        if marker in self.marker_list:
            self.marker_list.remove(marker)

        try:
            self.route_markers = [m for m in self.route_markers if m is not marker]
            self._rebuild_route_from_markers()
        except Exception:
            pass

        self.hide_info_card()
        self.current_selected_marker = None
        self.save_pins()
        self._trigger_route_metrics_recalc(force_routing=False)

    def clear_markers(self):
        to_delete = [m for m in self.marker_list if not getattr(m, "is_system", False)]
        for m in to_delete:
            try:
                m.delete()
            except Exception:
                pass
            try:
                self.marker_list.remove(m)
            except Exception:
                pass

        self.clear_search_marker()

        self.clear_route()
        self.hide_info_card()

        try:
            save_pin_records(self.pins_file, [])
        except OSError:
            logger.exception("Pins could not be cleared on disk.")
            messagebox.showerror("Speichern", "Pins konnten nicht gelöscht werden.")

        self._refresh_all_markers()

    # ---------- Search ----------
    def search_location(self):
        address = self.search_entry.get().strip()
        if not address:
            return
        self._search_job += 1
        job_id = self._search_job
        normalized = self._normalize_address_query(address)
        self.clear_search_marker()

        def _worker():
            try:
                latlng = self.geocoding_service.lookup(normalized) or self.geocoding_service.lookup(address)
                self.geocoding_service.save_cache()
            except OSError as exc:
                message = f"Cache konnte nicht gespeichert werden: {exc}"
                logger.exception("Geocode cache could not be saved.")
                self.after(0, lambda: self._on_search_error(job_id, message))
                return
            except Exception as exc:
                message = str(exc)
                logger.exception("Search geocoding failed for %s", address)
                self.after(0, lambda: self._on_search_error(job_id, message))
                return

            self.after(0, lambda: self._apply_search_result(job_id, address, latlng))

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_search_result(self, job_id: int, address: str, latlng):
        if job_id != self._search_job:
            return
        if not latlng:
            messagebox.showwarning("Suche", f"Ort '{address}' wurde nicht gefunden.")
            return

        lat, lng = latlng
        self.map_widget.set_position(lat, lng)
        self.map_widget.set_zoom(15)
        self.search_marker = self.map_widget.set_marker(
            lat,
            lng,
            text=f" Gesuchter Ort:\n {address} ",
            marker_color_outside="#3498db",
            marker_color_circle="white",
            command=self.on_marker_click,
        )
        self.search_marker.full_info = f" Gesuchter Ort:\n {address} "
        try:
            if hasattr(self.search_marker, "label"):
                self.search_marker.label.configure(fg_color="#2b2b2b", text_color="white")
        except (AttributeError, tk.TclError):
            logger.warning("Search marker label could not be styled.")

    def _on_search_error(self, job_id: int, message: str):
        if job_id != self._search_job:
            return
        messagebox.showerror("Fehler", f"Suche fehlgeschlagen: {message}")

    def clear_search_marker(self):
        if self.search_marker:
            try:
                self.search_marker.delete()
            except Exception:
                pass
            self.search_marker = None

    # ---------- Email ----------
    def open_email_client(self):
        marker = getattr(self, "current_selected_marker", None)
        if not marker:
            messagebox.showwarning("E-Mail", "Kein Pin ausgewählt.")
            return

        email = str(getattr(marker, "email", "")).strip()
        if not email or email.upper() == "N/A":
            messagebox.showwarning("E-Mail", "Für diesen Pin ist keine E-Mail-Adresse hinterlegt.")
            return

        auftrag = str(getattr(marker, "auftragsnummer", "")).strip()
        subject = f"Lieferung von Auftrag {auftrag}" if auftrag and auftrag.upper() != "N/A" else "Lieferung von Auftrag"
        subject_enc = quote(subject)

        try:
            webbrowser.open(f"mailto:{email}?subject={subject_enc}")
        except Exception as e:
            messagebox.showerror("E-Mail", f"E-Mail-Client konnte nicht geöffnet werden:\n{e}")

    def open_customer_editor(self, marker):
        """
        Popup zum Bearbeiten der Adress-/Kundendaten eines Pins.
        Aufrufbar aus Auftragsliste und Liefertouren (Doubleclick).
        """
        if not marker:
            return
        if getattr(marker, "is_system", False):
            messagebox.showwarning("Kundenkartei", "Depot/System-Pins können nicht bearbeitet werden.")
            return

        data = _normalize_order_and_delivery_addresses(dict(getattr(marker, "data", {}) or {}))
        # Sicherheit: Keys vorhanden
        for k in [
            "Name",
            "Strasse",
            "PLZ",
            "Ort",
            "Land",
            "Email",
            "Telefon",
            "Gewicht",
            "Produkte",
            "Auftragsnummer",
            "Bestelldatum",
            "Status",
            "Notizen",
            "Lieferart",
            "AuftragsadresseName",
            "AuftragsadressePLZ",
            "AuftragsadresseOrt",
            "LieferadresseName",
            "LieferadresseStrasse",
            "LieferadressePLZ",
            "LieferadresseOrt",
        ]:
            data.setdefault(k, "")
        data["Lieferart"] = _normalize_delivery_type(data.get("Lieferart"))

        dlg = ctk.CTkToplevel(self)
        dlg.title("Kundenkartei – Adresse bearbeiten")
        dlg.geometry("1320x980")
        dlg.resizable(True, True)
        dlg.minsize(700, 600)
        dlg.configure(fg_color=Theme.BG)
        dlg.attributes("-topmost", True)

        shell = ctk.CTkFrame(dlg, corner_radius=18, fg_color=Theme.PANEL, border_width=1, border_color=Theme.BORDER)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        shell.grid_columnconfigure(0, weight=1)

        # Header
        ctk.CTkLabel(shell, text="Kundenkartei", font=_font(18, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=(16, 6), sticky="w"
        )
        ctk.CTkLabel(shell, text=f"Auftragsnummer: {data.get('Auftragsnummer', '')}", font=_font(12),
                     text_color=Theme.SUBTEXT).grid(row=1, column=0, padx=16, pady=(0, 12), sticky="w")
        ctk.CTkLabel(
            shell,
            text=f"Auftragsadresse: {_format_order_address(data) or 'Nicht hinterlegt'}",
            font=_font(12),
            text_color=Theme.SUBTEXT,
        ).grid(row=2, column=0, padx=16, pady=(0, 12), sticky="w")

        form = ctk.CTkFrame(shell, fg_color="transparent")
        form.grid(row=3, column=0, padx=16, pady=(0, 10), sticky="nsew")
        form.grid_columnconfigure((0, 1), weight=1)

        def _row(r, label, value, col=0, colspan=1):
            ctk.CTkLabel(form, text=label, font=_font(12, "bold"), text_color=Theme.SUBTEXT).grid(
                row=r, column=col, padx=8, pady=(6, 2), sticky="w", columnspan=colspan
            )
            e = ctk.CTkEntry(form, height=36, corner_radius=12)
            e.grid(row=r + 1, column=col, padx=8, pady=(0, 6), sticky="ew", columnspan=colspan)
            e.insert(0, value if value is not None else "")
            return e

        # Felder (Adresse & Kontakt)
        ent_name = _row(0, "Liefername", data.get("LieferadresseName", data.get("Name", "")), col=0)
        ent_email = _row(0, "Email", data.get("Email", ""), col=1)
        ent_strasse = _row(2, "Lieferstrasse", data.get("LieferadresseStrasse", data.get("Strasse", "")), col=0, colspan=2)
        ent_plz = _row(4, "Liefer-PLZ", data.get("LieferadressePLZ", data.get("PLZ", "")), col=0)
        ent_ort = _row(4, "Lieferort", data.get("LieferadresseOrt", data.get("Ort", "")), col=1)
        ent_tel = _row(6, "Telefon", data.get("Telefon", ""), col=0)
        ent_gewicht = _row(6, "Gewicht", data.get("Gewicht", ""), col=1)

        # Status (optional editierbar – wenn du es NUR Adresse willst: Block einfach entfernen)
        ctk.CTkLabel(form, text="Status", font=_font(12, "bold"), text_color=Theme.SUBTEXT).grid(
            row=8, column=0, padx=8, pady=(10, 2), sticky="w", columnspan=2
        )
        status_menu = ctk.CTkOptionMenu(
            form,
            values=["nicht festgelegt", "Bestellt", "Auf dem Weg", "Im Lager"],
            corner_radius=12,
            height=36,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL,
            button_color=Theme.ACCENT,
            button_hover_color=Theme.ACCENT_HOVER,
            text_color=Theme.TEXT
        )
        status_menu.grid(row=9, column=0, padx=8, pady=(0, 6), sticky="ew", columnspan=2)
        status_menu.set(data.get("Status", "nicht festgelegt") or "nicht festgelegt")

        ctk.CTkLabel(form, text="Lieferart", font=_font(12, "bold"), text_color=Theme.SUBTEXT).grid(
            row=10, column=0, padx=8, pady=(10, 2), sticky="w", columnspan=2
        )
        delivery_type_menu = ctk.CTkOptionMenu(
            form,
            values=DELIVERY_TYPE_OPTIONS,
            corner_radius=12,
            height=36,
            font=_font(13, "bold"),
            fg_color=Theme.PANEL,
            button_color=Theme.ACCENT,
            button_hover_color=Theme.ACCENT_HOVER,
            text_color=Theme.TEXT,
        )
        delivery_type_menu.grid(row=11, column=0, padx=8, pady=(0, 6), sticky="ew", columnspan=2)
        delivery_type_menu.set(_normalize_delivery_type(data.get("Lieferart", DEFAULT_DELIVERY_TYPE)))

        ctk.CTkLabel(form, text="Notizen", font=_font(12, "bold"), text_color=Theme.SUBTEXT).grid(
            row=12, column=0, padx=8, pady=(10, 2), sticky="w", columnspan=2
        )

        notes_box = ctk.CTkTextbox(form, height=120, corner_radius=12)
        notes_box.grid(row=13, column=0, padx=8, pady=(0, 6), sticky="nsew", columnspan=2)
        notes_box.insert("1.0", data.get("Notizen", "") or "")
        form.grid_rowconfigure(13, weight=1)

        # Buttons unterhalb der Daten: "Auf Karte anzeigen" & "Tour anzeigen"
        action_row = ctk.CTkFrame(shell, fg_color="transparent")
        action_row.grid(row=4, column=0, padx=16, pady=(4, 8), sticky="ew")
        action_row.grid_columnconfigure((0, 1), weight=1)

        def _show_on_map():
            try:
                self.show_page("map")
                lat, lng = marker.position[0], marker.position[1]
                self.map_widget.set_position(lat, lng)
                self.map_widget.set_zoom(15)
                self.on_marker_click(marker)
            except Exception:
                pass

        def _show_tour():
            t = self._pin_used_in_any_tour(marker)
            if not t:
                messagebox.showinfo("Tour", "Dieser Pin ist aktuell in keiner Liefertour eingeplant.")
                return
            self.apply_tour(t)

        btn_map = ctk.CTkButton(
            action_row, text="Auf Karte anzeigen",
            height=40, corner_radius=14, font=_font(13, "bold"),
            fg_color=Theme.PANEL_2, hover_color=Theme.BORDER, text_color=Theme.TEXT,
            command=_show_on_map
        )
        btn_map.grid(row=0, column=0, padx=(0, 8), pady=6, sticky="ew")

        btn_tour = ctk.CTkButton(
            action_row, text="Tour anzeigen",
            height=40, corner_radius=14, font=_font(13, "bold"),
            fg_color=Theme.PANEL_2, hover_color=Theme.BORDER, text_color=Theme.TEXT,
            command=_show_tour
        )
        btn_tour.grid(row=0, column=1, padx=(8, 0), pady=6, sticky="ew")

        # deaktivieren wenn keine Tour
        if self._pin_used_in_any_tour(marker) is None:
            try:
                btn_tour.configure(state="disabled")
            except Exception:
                pass

        # Footer Buttons: Speichern / Schließen
        footer = ctk.CTkFrame(shell, fg_color="transparent")
        footer.grid(row=5, column=0, padx=16, pady=(8, 16), sticky="ew")
        footer.grid_columnconfigure((0, 1), weight=1)

        def _save():
            # Daten schreiben
            new_data = dict(getattr(marker, "data", {}) or {})
            new_data["Name"] = ent_name.get().strip()
            new_data["Email"] = ent_email.get().strip()
            new_data["Strasse"] = ent_strasse.get().strip()
            new_data["PLZ"] = ent_plz.get().strip()
            new_data["Ort"] = ent_ort.get().strip()
            new_data["LieferadresseName"] = new_data["Name"]
            new_data["LieferadresseStrasse"] = new_data["Strasse"]
            new_data["LieferadressePLZ"] = new_data["PLZ"]
            new_data["LieferadresseOrt"] = new_data["Ort"]
            new_data["Telefon"] = ent_tel.get().strip()
            new_data["Gewicht"] = ent_gewicht.get().strip()
            new_data["Status"] = status_menu.get().strip() or "nicht festgelegt"
            new_data["Lieferart"] = _normalize_delivery_type(delivery_type_menu.get().strip())
            new_data["Notizen"] = notes_box.get("1.0", "end-1c").strip()
            new_data = _normalize_order_and_delivery_addresses(new_data)

            marker.data = new_data
            marker.email = new_data.get("Email", "")
            # Status ggf. aktualisieren (Icon/Anzeigen)
            marker.status = new_data.get("Status", "nicht festgelegt")
            marker.full_info = self._build_popup_text(new_data)

            # UI/Files aktualisieren
            self.save_pins()

            # Marker-Icons neu setzen (wenn Status oder Tour-Highlight relevant)
            self._refresh_all_markers()

            # Info-Card updaten, falls gerade ausgewählt
            if self.current_selected_marker and self._marker_key(self.current_selected_marker) == self._marker_key(
                    marker):
                try:
                    self.on_marker_click(self.current_selected_marker)
                except Exception:
                    pass

            self.show_toast("Kundendaten gespeichert")
            try:
                dlg.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            footer, text="Schließen",
            height=40, corner_radius=14,
            fg_color=Theme.MUTED_BTN, hover_color=Theme.MUTED_BTN_HOVER,
            text_color=("white", "white"),
            command=dlg.destroy
        ).grid(row=0, column=0, padx=(0, 8), pady=6, sticky="ew")

        ctk.CTkButton(
            footer, text="Speichern",
            height=40, corner_radius=14, font=_font(13, "bold"),
            fg_color=Theme.SUCCESS, hover_color=Theme.SUCCESS_HOVER,
            command=_save
        ).grid(row=0, column=1, padx=(8, 0), pady=6, sticky="ew")

        dlg.grab_set()
        dlg.focus_force()

    # ---------- Config ----------
    def load_config(self):
        try:
            cfg = self.settings_manager.load()
            self.sql_data_dir = str(cfg.get("sql_data_dir") or self.sql_data_dir or DEFAULT_SQL_DATA_DIR).strip()
            self.sql_server_instance = str(cfg.get("sql_server_instance") or self.sql_server_instance or r".\SQLEXPRESS").strip()
            self.sql_database = str(cfg.get("sql_database") or "").strip() or infer_database_name_from_data_dir(self.sql_data_dir)
            self.appearance_preference = str(cfg.get("appearance_mode") or "System").title()
            self.quick_access_items = self.normalize_quick_access_items(cfg.get("quick_access_items", DEFAULT_QUICK_ACCESS_ITEMS))
            self.set_appearance_preference(self.appearance_preference, persist=False)
            self.refresh_quick_access_tools()
        except (OSError, ValueError):
            logger.exception("Configuration could not be loaded.")

    def save_config(self):
        try:
            self.settings_manager.save(
                {
                    "sql_data_dir": self.sql_data_dir or DEFAULT_SQL_DATA_DIR,
                    "sql_server_instance": self.sql_server_instance or r".\SQLEXPRESS",
                    "sql_database": self.sql_database or "",
                    "appearance_mode": self.get_appearance_preference(),
                    "quick_access_items": self.normalize_quick_access_items(self.quick_access_items),
                }
            )
        except (OSError, ValueError):
            logger.exception("Configuration could not be saved.")

    def _check_auto_backup_due(self):
        if self._auto_backup_running:
            return

        try:
            settings = self.settings_manager.load()
        except Exception:
            return

        if not settings.get("backups_enabled") or not settings.get("auto_backup_enabled"):
            return

        try:
            interval_days = int(settings.get("auto_backup_interval_days", 7))
        except Exception:
            interval_days = 7
        if interval_days < 1:
            interval_days = 1

        last_backup_iso = str(settings.get("last_backup_iso") or "").strip()
        due = True
        if last_backup_iso:
            try:
                last_backup_dt = datetime.fromisoformat(last_backup_iso.replace("Z", "+00:00"))
                due = (datetime.now(timezone.utc) - last_backup_dt.astimezone(timezone.utc)) >= timedelta(days=interval_days)
            except Exception:
                due = True

        if not due:
            return

        backup_dir = Path(str(settings.get("backup_dir") or self.settings_manager.default_backup_dir())).expanduser()
        mode = str(settings.get("backup_mode_default") or "full").strip().lower() or "full"
        self._auto_backup_running = True

        def _worker():
            try:
                manager = BackupManager(
                    app_name=self.title().strip() or "GAWELA Tourenplaner",
                    config_dir=Path(self.config_dir),
                    data_dir=Path(self.data_dir),
                    backup_dir=backup_dir,
                )
                backup_path = manager.create_backup(mode)
                manager.cleanup_old_backups(settings.get("backup_retention_days", 30))
                self.settings_manager.save({"last_backup_iso": datetime.now(timezone.utc).isoformat()})
                self.after(0, lambda: self._on_auto_backup_success(backup_path))
            except Exception:
                logger.exception("Automatic backup failed.")
                self.after(0, self._on_auto_backup_finished)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_auto_backup_success(self, _backup_path):
        self._auto_backup_running = False
        try:
            settings_page = self.pages.get("settings")
            if settings_page and hasattr(settings_page, "refresh"):
                settings_page.refresh()
        except AttributeError:
            logger.warning("Settings page refresh is not available after auto backup.")

    def _on_auto_backup_finished(self):
        self._auto_backup_running = False

    def select_sql_data_dir(self):
        folder = filedialog.askdirectory(title="SQL-Datenordner auswählen")
        if not folder:
            return
        self.sql_data_dir = folder
        guessed_database = infer_database_name_from_data_dir(self.sql_data_dir)
        if guessed_database:
            self.sql_database = guessed_database
        self.save_config()
        messagebox.showinfo(
            "Ordner gespeichert",
            f"SQL-Datenordner:\n{folder}\n\nDatenbank: {self.sql_database or 'nicht erkannt'}",
        )

    def _normalize_address_query(self, address: str, country_hint: str = "Schweiz") -> str:
        address = (address or "").strip()
        if not address:
            return ""
        lowered = address.lower()
        if country_hint and country_hint.lower() not in lowered:
            return f"{address}, {country_hint}"
        return address

    def _invalidate_tour_data_caches(self):
        self._tour_data_revision += 1
        self._calendar_payload_cache = {}
        self._calendar_payload_revision = -1
        self._tour_usage_cache = {}
        self._tour_usage_cache_revision = -1

    def get_calendar_payload_revision(self) -> int:
        return self._calendar_payload_revision

    def get_calendar_payload_map(self) -> dict:
        if self._calendar_payload_revision != self._tour_data_revision:
            payload = {}
            for tour in self._load_tours():
                date_key = _normalize_date_string(tour.get("date"))
                if not date_key:
                    continue
                entry = payload.setdefault(date_key, {"tours": 0, "assignments": 0, "titles": []})
                entry["tours"] += 1
                entry["assignments"] += tour_assignment_count(tour)
                title = self._tour_display_text(tour)
                if title not in entry["titles"]:
                    entry["titles"].append(title)
            self._calendar_payload_cache = payload
            self._calendar_payload_revision = self._tour_data_revision
        return self._calendar_payload_cache

    def _refresh_pages(self, *page_names: str, visible_only: bool = True):
        current_page = getattr(self, "current_page", None)
        for page_name in page_names:
            if visible_only and current_page != page_name:
                continue
            try:
                page = self.pages.get(page_name)
                if page and hasattr(page, "refresh"):
                    page.refresh()
            except Exception:
                pass

    def is_sql_import_running(self) -> bool:
        running = bool(getattr(self, "_sql_import_running", False))
        if not running:
            return False
        started_at = getattr(self, "_sql_import_started_at", None)
        if started_at is None:
            # Recover from stale state after crashes/exceptions where running stayed True.
            self._set_sql_import_status("Bereit", running=False)
            return False
        return True

    def get_sql_import_status_text(self) -> str:
        return str(getattr(self, "_sql_import_status_text", "Bereit") or "Bereit")

    def _set_sql_import_status(self, text: str, *, running: bool | None = None):
        self._sql_import_status_text = str(text or "").strip() or "Bereit"
        if running is not None:
            self._sql_import_running = bool(running)
        self._refresh_pages("list", "nonmap", visible_only=False)

    def _load_pending_sql_orders(self):
        self.pending_sql_orders = []
        try:
            payload = load_json_file(
                self.pending_orders_file,
                default=list,
                create_if_missing=False,
                backup_invalid=True,
            )
        except InvalidJsonFileError:
            logger.exception("Pending SQL orders file is invalid.")
            return
        except OSError:
            logger.exception("Pending SQL orders file could not be loaded.")
            return
        if not isinstance(payload, list):
            return
        items = []
        for raw in payload:
            if not isinstance(raw, dict):
                continue
            order_number = str(raw.get("Auftragsnummer") or "").strip()
            import_id = str(raw.get("ImportID") or "").strip()
            if not order_number and not import_id:
                continue
            item = dict(raw)
            item["Auftragsnummer"] = order_number
            item["ImportID"] = import_id
            item.setdefault("Status", "nicht festgelegt")
            item.setdefault("Lieferart", DEFAULT_DELIVERY_TYPE)
            item = _normalize_order_and_delivery_addresses(item)
            items.append(item)
        self.pending_sql_orders = items

    def _save_pending_sql_orders(self):
        try:
            atomic_write_json(self.pending_orders_file, list(self.pending_sql_orders or []))
        except OSError:
            logger.exception("Pending SQL orders could not be saved.")

    def _load_non_map_sql_orders(self):
        self.non_map_sql_orders = []
        try:
            payload = load_json_file(
                self.non_map_orders_file,
                default=list,
                create_if_missing=False,
                backup_invalid=True,
            )
        except InvalidJsonFileError:
            logger.exception("Non-map SQL orders file is invalid.")
            return
        except OSError:
            logger.exception("Non-map SQL orders file could not be loaded.")
            return
        if not isinstance(payload, list):
            return
        items = []
        for raw in payload:
            if not isinstance(raw, dict):
                continue
            order_number = str(raw.get("Auftragsnummer") or "").strip()
            import_id = str(raw.get("ImportID") or "").strip()
            if not order_number and not import_id:
                continue
            item = _normalize_order_and_delivery_addresses(dict(raw))
            item["Auftragsnummer"] = order_number
            item["ImportID"] = import_id
            item["NichtKarteKategorie"] = _normalize_non_map_order_category(
                item.get("NichtKarteKategorie") or item.get("Lieferart")
            )
            items.append(item)
        self.non_map_sql_orders = items

    def _save_non_map_sql_orders(self):
        try:
            atomic_write_json(self.non_map_orders_file, list(self.non_map_sql_orders or []))
        except OSError:
            logger.exception("Non-map SQL orders could not be saved.")

    def _write_geocode_failure_report(self, samples: list[dict]) -> str | None:
        if not samples:
            return None
        try:
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            path = Path(self.logs_dir) / f"sql_import_geocode_failed_{ts}.txt"
            lines = [
                "SQL-Import: Geocoding fehlgeschlagen",
                f"Erstellt: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}",
                f"Eintraege: {len(samples)}",
                "",
            ]
            for idx, item in enumerate(samples, start=1):
                order_number = str(item.get("Auftragsnummer") or "").strip()
                import_id = str(item.get("ImportID") or "").strip()
                name = str(item.get("Name") or "").strip()
                street = str(item.get("Strasse") or "").strip()
                plz = str(item.get("PLZ") or "").strip()
                ort = str(item.get("Ort") or "").strip()
                land = str(item.get("Land") or "").strip()
                query_list = item.get("queries") or []
                lines.append(
                    f"{idx}. ImportID={import_id} | Auftrag={order_number} | Name={name} | Adresse={street}, {plz} {ort}, {land}"
                )
                if query_list:
                    lines.append("   Versuche: " + " || ".join(str(q).strip() for q in query_list if str(q).strip()))
                lines.append("")
            path.write_text("\n".join(lines), encoding="utf-8")
            return str(path)
        except OSError:
            logger.exception("Geocoding failure report could not be written.")
            return None

    # ---------- SQL Import ----------
    def _build_geocode_query_candidates(self, payload: dict, country_hint: str) -> list[str]:
        normalized = _normalize_order_and_delivery_addresses(payload)
        street = str(normalized.get("LieferadresseStrasse") or "").strip()
        plz = str(normalized.get("LieferadressePLZ") or "").strip()
        ort = str(normalized.get("LieferadresseOrt") or "").strip()
        candidates = []
        if street or plz or ort:
            candidates.append(self._normalize_address_query(f"{street}, {plz} {ort}", country_hint=country_hint))
        if street and (plz or ort):
            candidates.append(self._normalize_address_query(f"{street} {plz} {ort}", country_hint=country_hint))
        if plz or ort:
            candidates.append(self._normalize_address_query(f"{plz} {ort}", country_hint=country_hint))
        if ort:
            candidates.append(self._normalize_address_query(ort, country_hint=country_hint))
        deduped = []
        seen = set()
        for text in candidates:
            q = str(text or "").strip()
            if not q:
                continue
            key = q.casefold()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(q)
        return deduped

    def _lookup_geocode_for_payload(self, payload: dict) -> tuple[tuple[float, float] | None, list[str], bool]:
        attempted_queries = []
        had_exception = False
        country_hint = _country_hint_from_code(payload.get("Land"))
        queries = self._build_geocode_query_candidates(payload, country_hint)
        latlng = None
        for query in queries:
            attempted_queries.append(query)
            try:
                latlng = self.geocoding_service.lookup(query)
            except Exception:
                logger.exception("Geocoding failed for order %s", str(payload.get("Auftragsnummer") or "").strip())
                had_exception = True
                latlng = None
            if latlng:
                break
        return latlng, attempted_queries, had_exception

    def import_from_sql(self, silent: bool = False):
        if self.is_sql_import_running():
            if not silent:
                messagebox.showinfo("Import", "Ein SQL-Import läuft bereits. Bitte warten, bis er abgeschlossen ist.")
            return

        self._import_job += 1
        job_id = self._import_job
        self._sql_import_started_at = datetime.now()
        self._set_sql_import_status(f"Läuft seit {self._sql_import_started_at.strftime('%H:%M:%S')}", running=True)

        existing_item_state = {}
        for marker in self.marker_list:
            if getattr(marker, "is_system", False):
                continue
            marker_data = dict(getattr(marker, "data", {}) or {})
            identity_key = self._item_identity_key(marker_data)
            if identity_key is None:
                continue
            try:
                lat, lng = marker.position[0], marker.position[1]
            except Exception:
                lat, lng = None, None
            existing_item_state[identity_key] = {
                "lat": lat,
                "lng": lng,
                "status": str(getattr(marker, "status", "nicht festgelegt") or "nicht festgelegt"),
                "has_coords": lat is not None and lng is not None,
            }
        existing_pending_state = {}
        for item in list(self.pending_sql_orders or []):
            if not isinstance(item, dict):
                continue
            identity_key = self._item_identity_key(item)
            if identity_key is None:
                continue
            existing_pending_state[identity_key] = dict(item)
        existing_non_map_state = {}
        for item in list(self.non_map_sql_orders or []):
            if not isinstance(item, dict):
                continue
            identity_key = self._item_identity_key(item)
            if identity_key is None:
                continue
            existing_non_map_state[identity_key] = dict(item)

        def _worker():
            upsert_items = []
            pending_items = []
            non_map_items = []
            error_text = None
            open_items_from_sql = set()
            open_map_items_from_sql = set()
            stats = {
                "rows_total": 0,
                "duplicate_sql_rows": 0,
                "skipped_non_map_delivery": 0,
                "created": 0,
                "updated": 0,
                "removed": 0,
                "removed_without_order_number": 0,
                "pending_total": 0,
                "pending_existing": 0,
                "geocoded_in_background": 0,
                "geocode_failed": 0,
                "geocode_errors": 0,
                "geocode_failed_samples": [],
            }
            try:
                database = str(self.sql_database or "").strip() or infer_database_name_from_data_dir(self.sql_data_dir)
                if not database:
                    raise RuntimeError("Keine Nutzdatenbank im SQL-Datenordner gefunden.")
                self.sql_database = database
                rows = fetch_order_rows(
                    server_instance=self.sql_server_instance,
                    database=database,
                    limit=20000,
                )
                stats["rows_total"] = len(rows)
                for row in rows:
                    order_number = str(row.get("Auftragsnummer") or "").strip()
                    if not order_number:
                        continue
                    import_id = str(row.get("ImportID") or "").strip()
                    identity_key = self._item_identity_key({"ImportID": import_id, "Auftragsnummer": order_number})
                    if identity_key is None:
                        continue
                    if identity_key in open_items_from_sql:
                        stats["duplicate_sql_rows"] += 1
                        continue
                    delivery_label, include_on_map = _resolve_sql_delivery_handling(row.get("Liefercode"))
                    open_items_from_sql.add(identity_key)
                    non_map_category = _normalize_non_map_order_category(delivery_label or row.get("Liefercode"))
                    base_payload = {
                        "ImportID": import_id,
                        "Auftragsnummer": order_number,
                        "Bestelldatum": _display_date_string(row.get("Bestelldatum")) or "N/A",
                        "Name": row.get("Name") or "Unbekannt",
                        "Strasse": row.get("Strasse") or "",
                        "PLZ": row.get("PLZ") or "",
                        "Ort": row.get("Ort") or "",
                        "AuftragsadresseName": row.get("AuftragName") or "",
                        "AuftragsadressePLZ": row.get("AuftragPLZ") or "",
                        "AuftragsadresseOrt": row.get("AuftragOrt") or "",
                        "LieferadresseName": row.get("Name") or "Unbekannt",
                        "LieferadresseStrasse": row.get("Strasse") or "",
                        "LieferadressePLZ": row.get("PLZ") or "",
                        "LieferadresseOrt": row.get("Ort") or "",
                        "Land": row.get("Land") or "",
                        "Email": row.get("Email") or "N/A",
                        "Telefon": row.get("Telefon") or "N/A",
                        "Gewicht": row.get("ProduktgewichtTotal") or row.get("Gewicht") or "N/A",
                        "Auftragsgewicht": row.get("Gewicht") or "",
                        "Produkte": row.get("Produkte") or "",
                        "Notizen": row.get("Notizen") or "",
                        "Lieferart": delivery_label,
                        "NichtKarteKategorie": non_map_category,
                    }
                    payload = _normalize_order_and_delivery_addresses(base_payload)
                    previous = existing_item_state.get(identity_key)
                    previous_pending = existing_pending_state.get(identity_key)
                    previous_non_map = existing_non_map_state.get(identity_key)
                    payload["Status"] = (previous or previous_pending or previous_non_map or {}).get("status", "nicht festgelegt")
                    if not include_on_map:
                        stats["skipped_non_map_delivery"] += 1
                        non_map_items.append(payload)
                        continue
                    open_map_items_from_sql.add(identity_key)
                    if previous and previous.get("has_coords"):
                        payload["lat"], payload["lng"] = previous["lat"], previous["lng"]
                        upsert_items.append(payload)
                        stats["updated"] += 1
                        continue
                    pending_items.append(payload)
                    if previous_pending:
                        stats["pending_existing"] += 1
                    else:
                        stats["created"] += 1
                stats["pending_total"] = len(pending_items)

                existing_map_identities = set(existing_item_state.keys()).union(existing_pending_state.keys())
                items_to_remove = sorted(existing_map_identities - open_map_items_from_sql)
                stats["removed"] = len(items_to_remove)
            except Exception as exc:
                logger.exception("SQL import failed.")
                error_text = str(exc)
                items_to_remove = []

            cache_error = None
            try:
                self.geocoding_service.save_cache()
            except OSError as exc:
                logger.exception("Geocode cache could not be saved after SQL import.")
                cache_error = str(exc)

            self.after(
                0,
                lambda: self._finish_sql_import(
                    job_id,
                    upsert_items,
                    pending_items,
                    non_map_items,
                    items_to_remove,
                    error_text,
                    cache_error,
                    silent,
                    stats,
                ),
            )

        threading.Thread(target=_worker, daemon=True).start()

    def _finish_sql_import(
        self,
        job_id: int,
        upsert_items: list[dict],
        pending_items: list[dict],
        non_map_items: list[dict],
        items_to_remove: list[tuple[str, str]],
        error_text: str | None,
        cache_error: str | None,
        silent: bool,
        stats: dict | None = None,
    ):
        if job_id != self._import_job:
            return
        try:
            marker_by_identity, markers_without_identity = self._collect_non_system_markers_by_identity()

            removed_markers = set()
            for marker in markers_without_identity:
                removed_markers.add(marker)
                self._delete_marker_safely(marker, remove_from_list=True)

            for identity_key in items_to_remove:
                marker = marker_by_identity.get(identity_key)
                if marker is None:
                    continue
                removed_markers.add(marker)
                self._delete_marker_safely(marker, remove_from_list=True)

            for item in upsert_items:
                identity_key = self._item_identity_key(item)
                marker = marker_by_identity.get(identity_key)
                if marker is None:
                    self._create_marker(item["lat"], item["lng"], item, item["Status"])
                    continue
                self._apply_payload_to_existing_marker(marker, item)

            pending_map = {self._item_identity_key(item): dict(item) for item in pending_items if self._item_identity_key(item)}
            non_map_map = {self._item_identity_key(item): dict(item) for item in non_map_items if self._item_identity_key(item)}
            for identity_key in items_to_remove:
                pending_map.pop(identity_key, None)
                non_map_map.pop(identity_key, None)
            self.pending_sql_orders = list(pending_map.values())
            self._save_pending_sql_orders()
            self.non_map_sql_orders = list(non_map_map.values())
            self._save_non_map_sql_orders()

            if removed_markers:
                stats = stats or {}
                stats["removed_without_order_number"] = len(markers_without_identity)
                self.route_markers = [m for m in self.route_markers if m not in removed_markers]
                try:
                    self._rebuild_route_from_markers()
                except Exception:
                    pass
                if getattr(self, "current_selected_marker", None) in removed_markers:
                    self.current_selected_marker = None
                    self.hide_info_card()

            self._refresh_after_sql_marker_changes(bool(upsert_items or removed_markers), show_map=True)
            self._refresh_pages("nonmap", visible_only=False)
            self._warn_sql_cache_save_error(cache_error)

            stats = stats or {}
            if error_text:
                self._fail_sql_import(error_text, message="SQL-Import konnte nicht verarbeitet werden")
                return

            pending_total = int(stats.get("pending_total", 0))
            if pending_total <= 0:
                finished_status = (
                    f"Fertig: gelesen {int(stats.get('rows_total', 0))}, neu {int(stats.get('created', 0))}, "
                    f"aktualisiert {int(stats.get('updated', 0))}"
                )
                self._set_sql_import_status(finished_status, running=False)
                if not silent:
                    messagebox.showinfo(
                        "Import",
                        "SQL-Import abgeschlossen.\n"
                        f"Datenbank: {self.sql_database or 'Unbekannt'}\n"
                        f"SQL-Aufträge gelesen: {int(stats.get('rows_total', 0))}\n"
                        f"Neu angelegt: {int(stats.get('created', 0))}\n"
                        f"Aktualisiert: {int(stats.get('updated', 0))}\n"
                        f"Ohne Karte (Lieferart): {int(stats.get('skipped_non_map_delivery', 0))}\n"
                        f"Entfernt (nicht mehr offen): {int(stats.get('removed', 0))}\n"
                        f"Entfernt (ohne Auftragsnummer): {int(stats.get('removed_without_order_number', 0))}\n"
                        f"Doppelte SQL-Zeilen: {int(stats.get('duplicate_sql_rows', 0))}",
                    )
                return

            self._set_sql_import_status(f"Liste bereit. Geocoding läuft 0/{pending_total}", running=True)
            self._pending_geocode_job += 1
            geocode_job_id = self._pending_geocode_job
            pending_snapshot = list(self.pending_sql_orders)
            threading.Thread(
                target=self._run_pending_geocode,
                args=(job_id, geocode_job_id, pending_snapshot, stats, silent),
                daemon=True,
            ).start()
        except Exception as exc:
            logger.exception("Finalizing SQL import failed.")
            self._fail_sql_import(exc, message="SQL-Import konnte nicht abgeschlossen werden")

    def _run_pending_geocode(
        self,
        import_job_id: int,
        geocode_job_id: int,
        pending_items: list[dict],
        phase1_stats: dict,
        silent: bool,
    ):
        try:
            resolved_items = []
            unresolved_items = []
            failed_samples = []
            geocode_errors = 0
            total = len(pending_items)
            for index, item in enumerate(pending_items, start=1):
                if import_job_id != self._import_job or geocode_job_id != self._pending_geocode_job:
                    return
                payload = dict(item)
                latlng, attempted_queries, had_exception = self._lookup_geocode_for_payload(payload)
                if had_exception:
                    geocode_errors += 1
                if latlng:
                    payload["lat"], payload["lng"] = latlng
                    resolved_items.append(payload)
                else:
                    unresolved_items.append(payload)
                    if len(failed_samples) < 30:
                        failed_samples.append(
                            {
                                "ImportID": payload.get("ImportID", ""),
                                "Auftragsnummer": payload.get("Auftragsnummer", ""),
                                "Name": payload.get("Name", ""),
                                "Strasse": payload.get("Strasse", ""),
                                "PLZ": payload.get("PLZ", ""),
                                "Ort": payload.get("Ort", ""),
                                "Land": payload.get("Land", ""),
                                "queries": list(attempted_queries),
                            }
                        )
                if index % 10 == 0 or index == total:
                    self.after(
                        0,
                        lambda i=index, t=total: self._set_sql_import_status(
                            f"Liste bereit. Geocoding läuft {i}/{t}",
                            running=True,
                        ),
                    )

            cache_error = None
            try:
                self.geocoding_service.save_cache()
            except OSError as exc:
                logger.exception("Geocode cache could not be saved after background geocoding.")
                cache_error = str(exc)

            self.after(
                0,
                lambda: self._finish_pending_geocode(
                    import_job_id,
                    geocode_job_id,
                    resolved_items,
                    unresolved_items,
                    failed_samples,
                    geocode_errors,
                    cache_error,
                    phase1_stats,
                    silent,
                ),
            )
        except Exception as exc:
            logger.exception("Pending geocoding worker crashed.")
            self.after(
                0,
                lambda: self._set_sql_import_status(f"Fehler: {exc}", running=False),
            )

    def _finish_pending_geocode(
        self,
        import_job_id: int,
        geocode_job_id: int,
        resolved_items: list[dict],
        unresolved_items: list[dict],
        failed_samples: list[dict],
        geocode_errors: int,
        cache_error: str | None,
        phase1_stats: dict,
        silent: bool,
    ):
        if import_job_id != self._import_job or geocode_job_id != self._pending_geocode_job:
            self._set_sql_import_status("Bereit", running=False)
            return
        try:
            marker_by_identity, _markers_without_identity = self._collect_non_system_markers_by_identity()

            for item in resolved_items:
                identity_key = self._item_identity_key(item)
                marker = marker_by_identity.get(identity_key)
                if marker is None:
                    self._create_marker(item["lat"], item["lng"], item, item.get("Status") or "nicht festgelegt")
                    continue
                self._apply_payload_to_existing_marker(marker, item)

            self.pending_sql_orders = list(unresolved_items or [])
            self._save_pending_sql_orders()

            self._refresh_after_sql_marker_changes(bool(resolved_items), show_map=False)
            self._warn_sql_cache_save_error(cache_error)

            report_path = self._write_geocode_failure_report(failed_samples or [])
            total_rows = int(phase1_stats.get("rows_total", 0))
            created = int(phase1_stats.get("created", 0))
            updated = int(phase1_stats.get("updated", 0))
            removed = int(phase1_stats.get("removed", 0))
            duplicates = int(phase1_stats.get("duplicate_sql_rows", 0))
            geocoded_background = len(resolved_items)
            geocode_failed = len(unresolved_items)

            finished_status = (
                f"Fertig: gelesen {total_rows}, neu {created}, aktualisiert {updated}, "
                f"geocodiert {geocoded_background}, offen ohne Koordinaten {geocode_failed}"
            )
            self._set_sql_import_status(finished_status, running=False)

            if not silent:
                preview_lines = []
                for item in (failed_samples or [])[:5]:
                    order_number = str(item.get("Auftragsnummer") or "").strip()
                    street = str(item.get("Strasse") or "").strip()
                    plz = str(item.get("PLZ") or "").strip()
                    ort = str(item.get("Ort") or "").strip()
                    land = str(item.get("Land") or "").strip()
                    preview_lines.append(f"- {order_number}: {street}, {plz} {ort}, {land}".strip())
                geocode_block = ""
                if preview_lines:
                    geocode_block = "\n\nBeispiele fehlgeschlagener Geocoding-Adressen:\n" + "\n".join(preview_lines)
                report_line = f"\nReport: {report_path}" if report_path else ""
                messagebox.showinfo(
                    "Import",
                    "SQL-Import (2 Phasen) abgeschlossen.\n"
                    f"Datenbank: {self.sql_database or 'Unbekannt'}\n"
                    f"SQL-Aufträge gelesen: {total_rows}\n"
                    f"Neu (in Liste): {created}\n"
                    f"Aktualisiert: {updated}\n"
                    f"Ohne Karte (Lieferart): {int(phase1_stats.get('skipped_non_map_delivery', 0))}\n"
                    f"Entfernt (nicht mehr offen): {removed}\n"
                    f"Doppelte SQL-Zeilen: {duplicates}\n"
                    f"Hintergrund-Geocoding erfolgreich: {geocoded_background}\n"
                    f"Ohne Koordinaten verbleibend: {geocode_failed}\n"
                    f"Geocoding-Fehler (Exceptions): {int(geocode_errors)}"
                    f"{report_line}"
                    f"{geocode_block}",
                )
        except Exception as exc:
            logger.exception("Finalizing pending geocoding failed.")
            self._fail_sql_import(exc, message="SQL-Import konnte nicht abgeschlossen werden")

    # ---------- Pins persistence ----------
    def save_pins(self):
        pins = []
        for m in self.marker_list:
            if getattr(m, "is_system", False):
                continue
            try:
                lat, lng = m.position[0], m.position[1]
                status = getattr(m, "status", "nicht festgelegt")
                data = _normalize_order_and_delivery_addresses(getattr(m, "data", {}) or {})
                data["Status"] = status
                m.data = data

                pins.append({
                    "lat": lat,
                    "lng": lng,
                    "status": status,
                    "data": data,
                })
            except (AttributeError, IndexError, TypeError):
                logger.warning("Skipping malformed marker during pin save.")

        try:
            save_pin_records(self.pins_file, pins)
        except OSError as e:
            logger.exception("Pins could not be saved.")
            messagebox.showerror("Speichern", f"Pins konnten nicht gespeichert werden:\n{e}")

    def load_pins(self):
        if not os.path.exists(self.pins_file):
            return
        if not hasattr(self, "map_widget") or self.map_widget is None:
            return

        try:
            pins = load_pin_records(self.pins_file)

            for m in list(self.marker_list):
                if getattr(m, "is_system", False):
                    continue
                try:
                    m.delete()
                except (AttributeError, tk.TclError):
                    logger.warning("Marker could not be deleted during pin reload.")
                try:
                    self.marker_list.remove(m)
                except ValueError:
                    logger.warning("Marker was already removed during pin reload.")

            for p in pins:
                lat = p.get("lat")
                lng = p.get("lng")
                status = p.get("status", "nicht festgelegt")
                data = p.get("data", {}) or {}
                data = _normalize_order_and_delivery_addresses(data)
                data["Status"] = status

                if lat is None or lng is None:
                    continue

                self._create_marker(lat, lng, data, status)

        except InvalidJsonFileError as e:
            logger.exception("Pins file is invalid.")
            messagebox.showerror("Laden", f"Pins-Datei ist beschädigt und wurde gesichert:\n{e}")
        except OSError as e:
            logger.exception("Pins could not be loaded.")
            messagebox.showerror("Laden", f"Pins konnten nicht geladen werden:\n{e}")

    # ---------- Close ----------
    def on_close(self):
        self.terminate_gps_native_window(timeout=2.0)
        self.save_pins()
        self.destroy()

    def _load_sidebar_icons(self) -> dict:
        icon_specs = {
            "Start": "Start.png",
            "Kalender": "Kalender.png",
            "Karte": "Karte & Suche.png",
            "GPS": "GPS.png",
            "Auftragsliste": "Auftragsliste.png",
            "Liefertouren": "Liefertouren.png",
            "Mitarbeiter": "Mitarbeiter.png",
            "Fahrzeuge": "Fahrzeuge.png",
            "Einstellungen": "Einstellungen.png",
        }
        loaded = {}
        for key, filename in icon_specs.items():
            path = os.path.join(self.sidebar_icons_dir, filename)
            try:
                image = Image.open(path)
                loaded[key] = ctk.CTkImage(light_image=image, dark_image=image, size=(22, 22))
            except Exception:
                loaded[key] = None
        return loaded

    def _clone_list(self, value):
        return copy.deepcopy(value if isinstance(value, list) else [])

    def _clone_dict(self, value, default=None):
        base = value if isinstance(value, dict) else (default if isinstance(default, dict) else {})
        return copy.deepcopy(base)

    def _load_employees(self) -> list:
        if self._employees_cache is not None:
            return self._clone_list(self._employees_cache)
        try:
            self._employees_cache = load_employees(self.employees_file)
        except Exception:
            self._employees_cache = []
        return self._clone_list(self._employees_cache)

    def _save_employees(self, employees: list) -> list:
        try:
            self._employees_cache = save_employees(self.employees_file, employees)
            return self._clone_list(self._employees_cache)
        except Exception as e:
            messagebox.showerror("Mitarbeiter", f"Konnte Mitarbeiter nicht speichern:\n{e}")
            return []

    def get_employee_map(self) -> dict:
        return {str(item.get("id")): item for item in self._load_employees() if isinstance(item, dict)}

    def get_employee_display_names(self, employee_ids) -> list:
        employee_map = self.get_employee_map()
        result = []
        for employee_id in employee_ids or []:
            key = str(employee_id).strip()
            if not key:
                continue
            employee = employee_map.get(key)
            if employee:
                short = str(employee.get("short") or "").strip()
                result.append(short or employee.get("name", key))
            else:
                result.append("Unbekannt")
        return result

    def format_employee_summary(self, employee_ids) -> str:
        names = self.get_employee_display_names(employee_ids)
        return ", ".join(names) if names else "Keine Mitarbeiter ausgewählt"

    def save_employee_record(self, employee: dict):
        name = str((employee or {}).get("name") or "").strip()
        if not name:
            raise ValueError("Der Name ist erforderlich.")

        employees = self._load_employees()
        employee_id = str(employee.get("id") or uuid4())
        stored = None
        for item in employees:
            if str(item.get("id")) == employee_id:
                stored = item
                break

        if stored is None:
            employees.append(employee)
        else:
            stored.update(employee)

        self._save_employees(employees)
        self.update_route_employee_summary()
        self._refresh_pages("employees", "tours", visible_only=True)

    def delete_employee_record(self, employee_id):
        key = str(employee_id or "").strip()
        employees = [item for item in self._load_employees() if str(item.get("id")) != key]
        self._save_employees(employees)
        self.update_route_employee_summary()
        self._refresh_pages("employees", "tours", visible_only=True)

    # ---------- Vehicles / Trailers ----------
    def _load_vehicle_data(self) -> dict:
        if self._vehicle_data_cache is not None:
            return self._clone_dict(self._vehicle_data_cache, {"vehicles": [], "trailers": []})
        try:
            self._vehicle_data_cache = load_vehicles(self.vehicles_file)
        except Exception:
            self._vehicle_data_cache = {"vehicles": [], "trailers": []}
        return self._clone_dict(self._vehicle_data_cache, {"vehicles": [], "trailers": []})

    def _save_vehicle_data(self, payload: dict) -> dict:
        try:
            self._vehicle_data_cache = save_vehicles(self.vehicles_file, payload)
            return self._clone_dict(self._vehicle_data_cache, {"vehicles": [], "trailers": []})
        except Exception as e:
            messagebox.showerror("Fahrzeuge", f"Konnte Fahrzeuge nicht speichern:\n{e}")
            return {"vehicles": [], "trailers": []}

    def get_vehicle_map(self, include_inactive: bool = True) -> dict:
        result = {}
        for item in self._load_vehicle_data().get("vehicles", []):
            if not include_inactive and not item.get("active", True):
                continue
            result[str(item.get("id"))] = item
        return result

    def get_trailer_map(self, include_inactive: bool = True) -> dict:
        result = {}
        for item in self._load_vehicle_data().get("trailers", []):
            if not include_inactive and not item.get("active", True):
                continue
            result[str(item.get("id"))] = item
        return result

    def get_active_vehicles(self) -> list:
        return list(self.get_vehicle_map(include_inactive=False).values())

    def get_active_trailers(self) -> list:
        return list(self.get_trailer_map(include_inactive=False).values())

    def _format_vehicle_label(self, vehicle: dict) -> str:
        if not isinstance(vehicle, dict):
            return "Bitte wählen"
        name = str(vehicle.get("name") or "").strip() or "Unbekanntes Fahrzeug"
        plate = str(vehicle.get("license_plate") or "").strip()
        return f"{name} ({plate})" if plate else name

    def _format_trailer_label(self, trailer: dict) -> str:
        if not isinstance(trailer, dict):
            return "Kein Anhänger"
        name = str(trailer.get("name") or "").strip() or "Unbekannter Anhänger"
        plate = str(trailer.get("license_plate") or "").strip()
        return f"{name} ({plate})" if plate else name

    def format_tour_vehicle_summary(self, tour: dict) -> str:
        tour = tour if isinstance(tour, dict) else {}
        vehicle = self.get_vehicle_map().get(str(tour.get("vehicle_id") or "").strip())
        trailer = self.get_trailer_map().get(str(tour.get("trailer_id") or "").strip())
        if vehicle is None:
            return "Kein Fahrzeug"
        summary = self._format_vehicle_label(vehicle)
        if trailer is not None:
            summary = f"{summary} + {self._format_trailer_label(trailer)}"
        return summary

    def build_vehicle_option_map(self, selected_id=None) -> tuple[list, dict]:
        options = {}
        selected_key = str(selected_id or "").strip()
        for vehicle in self.get_active_vehicles():
            label = self._format_vehicle_label(vehicle)
            options[label] = str(vehicle.get("id"))
        if selected_key and selected_key not in options.values():
            vehicle = self.get_vehicle_map().get(selected_key)
            if vehicle is not None:
                options[f"{self._format_vehicle_label(vehicle)} [inaktiv]"] = selected_key
        labels = list(options.keys()) or ["Keine aktiven Fahrzeuge"]
        return labels, options

    def build_trailer_option_map(self, selected_id=None) -> tuple[list, dict]:
        options = {"Kein Anhänger": None}
        selected_key = str(selected_id or "").strip()
        for trailer in self.get_active_trailers():
            label = self._format_trailer_label(trailer)
            options[label] = str(trailer.get("id"))
        if selected_key and selected_key not in [value for value in options.values() if value]:
            trailer = self.get_trailer_map().get(selected_key)
            if trailer is not None:
                options[f"{self._format_trailer_label(trailer)} [inaktiv]"] = selected_key
        return list(options.keys()), options

    def get_vehicle_option_label(self, vehicle_id, options: dict) -> str:
        key = str(vehicle_id or "").strip()
        for label, value in options.items():
            if str(value or "").strip() == key:
                return label
        labels, _ = self.build_vehicle_option_map(key)
        return labels[0]

    def get_trailer_option_label(self, trailer_id, options: dict) -> str:
        key = str(trailer_id or "").strip()
        if not key:
            return "Kein Anhänger"
        for label, value in options.items():
            if str(value or "").strip() == key:
                return label
        labels, _ = self.build_trailer_option_map(key)
        return labels[0]

    def upsert_vehicle_record(self, vehicle: dict):
        self._vehicle_data_cache = upsert_vehicle(self.vehicles_file, vehicle)
        self.update_route_resource_summary()
        self._refresh_pages("vehicles", "tours", visible_only=True)

    def upsert_trailer_record(self, trailer: dict):
        self._vehicle_data_cache = upsert_trailer(self.vehicles_file, trailer)
        self.update_route_resource_summary()
        self._refresh_pages("vehicles", "tours", visible_only=True)

    def delete_vehicle_record(self, vehicle_id, suppress_in_use_check: bool = False):
        key = str(vehicle_id or "").strip()
        if not suppress_in_use_check:
            for tour in self._load_tours():
                if str(tour.get("vehicle_id") or "").strip() == key:
                    raise ValueError("Fahrzeug ist in einer Tour hinterlegt. Bitte stattdessen deaktivieren.")
        self._vehicle_data_cache = delete_vehicle(self.vehicles_file, key)
        self.update_route_resource_summary()
        self._refresh_pages("vehicles", "tours", visible_only=True)

    def delete_trailer_record(self, trailer_id, suppress_in_use_check: bool = False):
        key = str(trailer_id or "").strip()
        if not suppress_in_use_check:
            for tour in self._load_tours():
                if str(tour.get("trailer_id") or "").strip() == key:
                    raise ValueError("Anhänger ist in einer Tour hinterlegt. Bitte stattdessen deaktivieren.")
        self._vehicle_data_cache = delete_trailer(self.vehicles_file, key)
        self.update_route_resource_summary()
        self._refresh_pages("vehicles", "tours", visible_only=True)

    def _find_tour_resource_conflicts(self, tour_date: str, vehicle_id: str, trailer_id=None, exclude_tour_id=None) -> list:
        conflicts = []
        date_key = _normalize_date_string(tour_date)
        vehicle_key = str(vehicle_id or "").strip()
        trailer_key = str(trailer_id or "").strip()

        if not date_key or not vehicle_key:
            return conflicts

        for other in self._load_tours():
            other_id = other.get("id")
            if exclude_tour_id is not None and other_id == exclude_tour_id:
                continue
            if _normalize_date_string(other.get("date")) != date_key:
                continue

            other_title = self._tour_display_text(other)
            if str(other.get("vehicle_id") or "").strip() == vehicle_key:
                conflicts.append(f"Fahrzeug bereits eingeplant: {other_title}")
            if trailer_key and str(other.get("trailer_id") or "").strip() == trailer_key:
                conflicts.append(f"Anhänger bereits eingeplant: {other_title}")
        return conflicts

    def _confirm_tour_resource_conflicts(self, conflicts: list) -> bool:
        if not conflicts:
            return True
        # Konfliktprüfung: Doppelbelegung am selben Tag wird standardmäßig nur gewarnt,
        # damit bestehende Workflows erhalten bleiben und der Benutzer bewusst übersteuern kann.
        detail = "\n".join(conflicts)
        return messagebox.askyesno(
            "Ressourcenkonflikt",
            f"{detail}\n\nTrotzdem speichern?",
        )

    def _validate_tour_vehicle_selection(self, vehicle_id: str, trailer_id=None):
        vehicle_key = str(vehicle_id or "").strip()
        trailer_key = str(trailer_id or "").strip()
        vehicle = self.get_vehicle_map().get(vehicle_key)
        trailer = self.get_trailer_map().get(trailer_key) if trailer_key else None

        if not vehicle_key:
            raise ValueError("Bitte ein Fahrzeug auswählen.")
        if vehicle is None:
            raise ValueError("Das gewählte Fahrzeug wurde nicht gefunden.")
        if not vehicle.get("active", True):
            raise ValueError("Das gewählte Fahrzeug ist inaktiv und kann nicht eingeplant werden.")
        if trailer is not None and not trailer.get("active", True):
            raise ValueError("Der gewählte Anhänger ist inaktiv und kann nicht eingeplant werden.")
        if trailer is not None and int(vehicle.get("max_trailer_load_kg", 0) or 0) <= 0:
            raise ValueError("Das gewählte Fahrzeug hat keine freigegebene Anhängelast.")

    def _resource_conflict_hint(self, tour_date: str, vehicle_id: str, trailer_id=None, exclude_tour_id=None) -> str:
        conflicts = self._find_tour_resource_conflicts(tour_date, vehicle_id, trailer_id, exclude_tour_id=exclude_tour_id)
        if not conflicts:
            return ""
        return " | ".join(conflicts)

    # ---------- Tours ----------
    def _load_tours(self) -> list:
        if self._tours_cache is not None:
            return self._clone_list(self._tours_cache)
        try:
            self._tours_cache = load_tours(self.tours_file)
        except Exception:
            self._tours_cache = []
        return self._clone_list(self._tours_cache)

    def _tour_display_text(self, t: dict) -> str:
        date = _normalize_date_string(t.get("date"))
        name = (t.get("name") or "").strip()
        if name and date:
            return f"{date} – {name}"
        if date:
            return date
        if name:
            return name
        return f"Tour {t.get('id', '')}".strip()

    def set_current_route_employee_ids(self, employee_ids):
        cleaned = []
        for value in employee_ids or []:
            text = str(value).strip()
            if text and text not in cleaned:
                cleaned.append(text)
        self.current_route_employee_ids = cleaned[:2]
        self.update_route_employee_summary()

    def update_route_employee_summary(self):
        label = getattr(self, "lbl_route_employee_summary", None)
        if label is None:
            return
        names = self.get_employee_display_names(self.current_route_employee_ids)
        label.configure(text=", ".join(names) if names else "Keine Mitarbeiter ausgewählt")

    def update_route_resource_summary(self):
        vehicle_label = getattr(self, "lbl_route_vehicle_summary", None)
        trailer_label = getattr(self, "lbl_route_trailer_summary", None)
        if vehicle_label is not None:
            vehicle = self.get_vehicle_map().get(str(self.current_route_vehicle_id or "").strip())
            vehicle_label.configure(text=self._format_vehicle_label(vehicle) if vehicle else "Kein Fahrzeug ausgewählt")
        if trailer_label is not None:
            trailer = self.get_trailer_map().get(str(self.current_route_trailer_id or "").strip())
            trailer_label.configure(text=self._format_trailer_label(trailer) if trailer else "Kein Anhänger")

    def set_current_route_resources(self, vehicle_id=None, trailer_id=None):
        self.current_route_vehicle_id = str(vehicle_id or "").strip() or None
        self.current_route_trailer_id = str(trailer_id or "").strip() or None
        self.update_route_resource_summary()

    def open_employee_picker(self, selected_ids=None, on_apply=None, title="Mitarbeiter wählen"):
        employees = self._load_employees()
        visible = [item for item in employees if item.get("active", True)]
        existing_ids = []
        for value in selected_ids or []:
            key = str(value).strip()
            if key and key not in existing_ids:
                existing_ids.append(key)

        for item in employees:
            key = str(item.get("id"))
            if key in existing_ids and all(str(v.get("id")) != key for v in visible):
                visible.append(item)

        if not visible:
            messagebox.showinfo("Mitarbeiter", "Es sind keine aktiven Mitarbeiter vorhanden.")
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title(title)
        dlg.geometry("520x520")
        dlg.resizable(True, True)
        dlg.configure(fg_color=Theme.BG)
        dlg.attributes("-topmost", True)

        shell = ctk.CTkFrame(dlg, corner_radius=18, fg_color=Theme.PANEL, border_width=1, border_color=Theme.BORDER)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(shell, text=title, font=_font(16, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=(16, 10), sticky="w"
        )

        scroll = ctk.CTkScrollableFrame(
            shell,
            corner_radius=14,
            fg_color=Theme.PANEL_2,
            **_scrollable_frame_kwargs(),
        )
        scroll.grid(row=1, column=0, padx=16, pady=(0, 12), sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        variables = {}
        info = ctk.CTkLabel(shell, text="Maximal 2 Mitarbeiter auswählbar.", font=_font(12), text_color=Theme.SUBTEXT)
        info.grid(row=2, column=0, padx=16, pady=(0, 8), sticky="w")

        def _sync_limit(changed_id=None):
            chosen = [emp_id for emp_id, var in variables.items() if bool(var.get())]
            if len(chosen) > 2 and changed_id in variables:
                variables[changed_id].set(False)
                messagebox.showwarning("Mitarbeiter", "Es können maximal 2 Mitarbeiter ausgewählt werden.")
                chosen = [emp_id for emp_id, var in variables.items() if bool(var.get())]
            info.configure(text=f"{len(chosen)} von 2 Mitarbeitern ausgewählt.")

        for row_idx, employee in enumerate(visible):
            employee_id = str(employee.get("id"))
            var = tk.BooleanVar(value=employee_id in existing_ids)
            variables[employee_id] = var

            display = employee.get("name", "")
            short = str(employee.get("short") or "").strip()
            if short:
                display = f"{display} ({short})"
            if not employee.get("active", True):
                display = f"{display} [inaktiv]"

            ctk.CTkCheckBox(
                scroll,
                text=display,
                variable=var,
                onvalue=True,
                offvalue=False,
                text_color=Theme.TEXT,
                command=lambda value=employee_id: _sync_limit(value),
            ).grid(row=row_idx, column=0, padx=10, pady=8, sticky="w")

        _sync_limit()

        btns = ctk.CTkFrame(shell, fg_color="transparent")
        btns.grid(row=3, column=0, padx=16, pady=(0, 16), sticky="ew")
        btns.grid_columnconfigure((0, 1), weight=1)

        def _apply():
            chosen = [emp_id for emp_id, var in variables.items() if bool(var.get())][:2]
            if callable(on_apply):
                on_apply(chosen)
            try:
                dlg.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            btns,
            text="Abbrechen",
            height=40,
            corner_radius=14,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=dlg.destroy,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            btns,
            text="Auswählen",
            height=40,
            corner_radius=14,
            font=_font(13, "bold"),
            fg_color=Theme.SUCCESS,
            hover_color=Theme.SUCCESS_HOVER,
            command=_apply,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        dlg.grab_set()
        dlg.focus_force()

    def open_route_resource_picker(self, selected_vehicle_id=None, selected_trailer_id=None, on_apply=None, title="Fahrzeug wählen"):
        vehicle_labels, vehicle_options = self.build_vehicle_option_map(selected_vehicle_id)
        trailer_labels, trailer_options = self.build_trailer_option_map(selected_trailer_id)

        dlg = ctk.CTkToplevel(self)
        dlg.title(title)
        dlg.geometry("560x420")
        dlg.resizable(True, True)
        dlg.configure(fg_color=Theme.BG)
        dlg.attributes("-topmost", True)

        shell = ctk.CTkFrame(dlg, corner_radius=18, fg_color=Theme.PANEL, border_width=1, border_color=Theme.BORDER)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        shell.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(shell, text=title, font=_font(16, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=(16, 10), sticky="w"
        )

        vehicle_var = tk.StringVar(value=self.get_vehicle_option_label(selected_vehicle_id, vehicle_options))
        trailer_var = tk.StringVar(value=self.get_trailer_option_label(selected_trailer_id, trailer_options))

        ctk.CTkLabel(shell, text="Fahrzeug *", font=_font(12, "bold"), text_color=Theme.TEXT).grid(
            row=1, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        ctk.CTkOptionMenu(shell, height=36, corner_radius=12, values=vehicle_labels, variable=vehicle_var).grid(
            row=2, column=0, padx=16, pady=(0, 10), sticky="ew"
        )

        ctk.CTkLabel(shell, text="Anhänger", font=_font(12, "bold"), text_color=Theme.TEXT).grid(
            row=3, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        ctk.CTkOptionMenu(shell, height=36, corner_radius=12, values=trailer_labels, variable=trailer_var).grid(
            row=4, column=0, padx=16, pady=(0, 12), sticky="ew"
        )

        btns = ctk.CTkFrame(shell, fg_color="transparent")
        btns.grid(row=5, column=0, padx=16, pady=(6, 16), sticky="ew")
        btns.grid_columnconfigure((0, 1), weight=1)

        def _apply():
            vehicle_id = vehicle_options.get(vehicle_var.get())
            trailer_id = trailer_options.get(trailer_var.get())
            if callable(on_apply):
                on_apply(vehicle_id, trailer_id)
            try:
                dlg.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            btns,
            text="Abbrechen",
            height=40,
            corner_radius=14,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=dlg.destroy,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            btns,
            text="Auswählen",
            height=40,
            corner_radius=14,
            font=_font(13, "bold"),
            fg_color=Theme.SUCCESS,
            hover_color=Theme.SUCCESS_HOVER,
            command=_apply,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        dlg.grab_set()
        dlg.focus_force()

    def open_tour_picker(self, anchor_widget):
        if getattr(self, "_tour_picker_popup", None) is not None:
            try:
                self._tour_picker_popup.destroy()
            except Exception:
                pass
            self._tour_picker_popup = None
            return

        tours = self._load_tours()
        if not tours:
            messagebox.showinfo("Tour", "Keine gespeicherten Liefertouren vorhanden.")
            return

        popup = ctk.CTkToplevel(self)
        self._tour_picker_popup = popup

        popup.title("Tour auswählen")
        popup.resizable(True, True)
        popup.attributes("-topmost", True)
        try:
            popup.attributes("-alpha", 0.96)
        except Exception:
            pass

        self.update_idletasks()
        x = anchor_widget.winfo_rootx()
        y = anchor_widget.winfo_rooty() + anchor_widget.winfo_height() + 8
        popup.geometry(f"520x520+{x}+{y}")

        popup.configure(fg_color=Theme.BG)
        popup.grid_columnconfigure(0, weight=1)
        popup.grid_rowconfigure(0, weight=1)

        def _close():
            popup_ref = getattr(self, "_tour_picker_popup", None)
            if popup_ref is not None:
                try:
                    if popup_ref.winfo_exists():
                        try:
                            popup_ref.grab_release()
                        except tk.TclError:
                            pass
                        popup_ref.withdraw()
                        popup_ref.after_idle(
                            lambda ref=popup_ref: ref.destroy() if ref.winfo_exists() else None
                        )
                except (AttributeError, tk.TclError):
                    pass
                self._tour_picker_popup = None

        def _apply_selected():
            tid = getattr(self, "_tour_picker_selected_tid", None)
            if tid is None:
                return

            chosen = None
            for t in tours:
                if t.get("id") == tid:
                    chosen = t
                    break
            if not chosen:
                return

            self.apply_tour(chosen)
            if hasattr(self, "lbl_current_tour"):
                self.lbl_current_tour.configure(text=self._tour_display_text(chosen))
            _close()

        glass = ctk.CTkFrame(
            popup,
            corner_radius=22,
            fg_color=Theme.OVERLAY_PANEL,
            border_width=1,
            border_color=Theme.OVERLAY_BORDER,
        )
        glass.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        glass.grid_columnconfigure(0, weight=1)
        glass.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            glass,
            text="Tour auswählen",
            font=_font(14, "bold"),
            text_color=Theme.TEXT
        ).grid(row=0, column=0, padx=18, pady=(16, 10), sticky="w")

        header = ctk.CTkFrame(glass, fg_color="transparent")
        header.grid(row=1, column=0, padx=18, pady=(0, 6), sticky="ew")
        header.grid_columnconfigure(0, weight=0)
        header.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(header, text="Datum", font=_font(12, "bold"), text_color=Theme.SUBTEXT).grid(
            row=0, column=0, sticky="w"
        )
        ctk.CTkLabel(header, text="Name", font=_font(12, "bold"), text_color=Theme.SUBTEXT).grid(
            row=0, column=1, sticky="w", padx=(14, 0)
        )

        list_shell = ctk.CTkFrame(
            glass,
            corner_radius=18,
            fg_color=Theme.OVERLAY_PANEL_2,
            border_width=1,
            border_color=Theme.OVERLAY_BORDER,
        )
        list_shell.grid(row=2, column=0, padx=18, pady=(0, 18), sticky="nsew")
        list_shell.grid_columnconfigure(0, weight=1)
        list_shell.grid_rowconfigure(0, weight=1)

        inner = ctk.CTkFrame(list_shell, corner_radius=14, fg_color="transparent")
        inner.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        inner.grid_columnconfigure(0, weight=1)
        inner.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(
            inner,
            corner_radius=14,
            fg_color=Theme.OVERLAY_PANEL_2,
            **_scrollable_frame_kwargs(),
        )
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        self._tour_picker_selected_tid = None
        row_widgets = {}

        def _set_selected(tid: int):
            self._tour_picker_selected_tid = tid
            for t_id, w in row_widgets.items():
                if t_id == tid:
                    w.configure(
                        fg_color=Theme.SELECTION,
                        border_color=Theme.ACCENT,
                        border_width=1
                    )
                else:
                    w.configure(
                        fg_color="transparent",
                        border_color=Theme.OVERLAY_BORDER,
                        border_width=1
                    )

        for t in tours:
            tid = t.get("id")
            try:
                tid_int = int(tid)
            except Exception:
                continue

            date = _normalize_date_string(t.get("date"))
            name = (t.get("name") or "").strip()

            row = ctk.CTkFrame(
                scroll,
                corner_radius=14,
                fg_color="transparent",
                border_width=1,
                border_color=Theme.OVERLAY_BORDER
            )
            row.grid(sticky="ew", padx=6, pady=6)
            row.grid_columnconfigure(0, weight=0)
            row.grid_columnconfigure(1, weight=1)

            lbl_date = ctk.CTkLabel(row, text=date, font=_font(12), text_color=Theme.TEXT)
            lbl_date.grid(row=0, column=0, sticky="w", padx=(12, 8), pady=10)

            lbl_name = ctk.CTkLabel(row, text=name, font=_font(12), text_color=Theme.TEXT)
            lbl_name.grid(row=0, column=1, sticky="w", padx=(14, 12), pady=10)

            row_widgets[tid_int] = row

            def _click_factory(tid_val: int):
                return lambda e=None: _set_selected(tid_val)

            def _dbl_factory(tid_val: int):
                def _do(e=None):
                    _set_selected(tid_val)
                    _apply_selected()
                return _do

            for w in (row, lbl_date, lbl_name):
                w.bind("<Button-1>", _click_factory(tid_int))
                w.bind("<Double-Button-1>", _dbl_factory(tid_int))

        popup.bind("<Escape>", lambda e: _close())
        popup.bind("<Return>", lambda e: _apply_selected())
        popup.protocol("WM_DELETE_WINDOW", _close)

        first = next((t for t in tours if isinstance(t, dict) and t.get("id") is not None), None)
        if first:
            try:
                _set_selected(int(first.get("id")))
            except Exception:
                pass

        popup.lift()
        popup.after_idle(lambda ref=popup: ref.focus_force() if ref.winfo_exists() else None)

    def _save_tours(self, tours: list):
        try:
            self._tours_cache = save_tours(self.tours_file, tours)
        except Exception as e:
            messagebox.showerror("Touren", f"Konnte Touren nicht speichern:\n{e}")
            return
        self._invalidate_tour_data_caches()
        self._refresh_tour_related_views()

    def _refresh_tour_related_views(self):
        try:
            page = self.pages.get("calendar")
            if page and hasattr(page, "refresh_calendar"):
                page.refresh_calendar()
        except Exception:
            pass

        self._refresh_pages("tours", visible_only=True)

    def save_current_tour(self, tour_date: str, tour_name: str = "", employee_ids=None, vehicle_id=None, trailer_id=None):
        tours = self._load_tours()
        existing_keys = set()
        self._sync_current_route_stop_data_from_markers()

        for t in tours:
            for s in t.get("stops", []):
                if not isinstance(s, dict):
                    continue
                if s.get("auftragsnummer"):
                    existing_keys.add(("auftrag", str(s.get("auftragsnummer")).strip()))
                elif s.get("lat") is not None and s.get("lng") is not None:
                    existing_keys.add(("coord", round(float(s["lat"]), 7), round(float(s["lng"]), 7)))

        for m in self.route_markers:
            if self._is_depot_marker(m):
                continue
            if self._marker_key(m) in existing_keys:
                messagebox.showwarning(
                    "Tour speichern",
                    "Mindestens ein Stopp ist bereits in einer anderen Liefertour vorhanden.\n"
                    "Speichern wurde abgebrochen."
                )
                return False

        normal_stops = [m for m in self.route_markers if m and not self._is_depot_marker(m)]
        if len(normal_stops) < 1:
            messagebox.showwarning("Tour speichern", "Für eine Liefertour brauchst du mindestens 1 Stopp (zusätzlich zum Depot).")
            return False

        tour_date = _normalize_date_string(tour_date)
        if not tour_date:
            messagebox.showwarning("Tour speichern", "Bitte ein Datum eintragen (z.B. 28-02-2026).")
            return False

        parsed_tour_date = parse_date(tour_date)
        if parsed_tour_date and parsed_tour_date < datetime.now().date():
            messagebox.showwarning("Tour speichern", "Das gewählte Datum liegt in der Vergangenheit.")

        tour_name = (tour_name or "").strip()
        selected_employee_ids = []
        for value in (self.current_route_employee_ids if employee_ids is None else employee_ids) or []:
            key = str(value).strip()
            if key and key not in selected_employee_ids:
                selected_employee_ids.append(key)

        if not (1 <= len(selected_employee_ids) <= 2):
            messagebox.showwarning("Tour speichern", "Bitte 1 oder 2 Mitarbeiter für die Liefertour auswählen.")
            return False

        selected_vehicle_id = str(self.current_route_vehicle_id if vehicle_id is None else vehicle_id or "").strip()
        selected_trailer_id = str(self.current_route_trailer_id if trailer_id is None else trailer_id or "").strip() or None
        try:
            self._validate_tour_vehicle_selection(selected_vehicle_id, selected_trailer_id)
        except ValueError as exc:
            messagebox.showwarning("Tour speichern", str(exc))
            return False

        conflicts = self._find_tour_resource_conflicts(tour_date, selected_vehicle_id, selected_trailer_id)
        if not self._confirm_tour_resource_conflicts(conflicts):
            return False

        stops = []
        stop_map = {self._marker_key(m): stop for m, stop in zip([m for m in self.route_markers if m and not self._is_depot_marker(m)], self.current_route_stop_data)}
        for order, m in enumerate(self.route_markers, start=1):
            if self._is_depot_marker(m):
                continue

            stop = dict(stop_map.get(self._marker_key(m), self._make_default_stop_from_marker(m, order=order)))
            stop["order"] = len(stops) + 1
            stops.append(stop)

        existing_ids = {t.get("id") for t in tours if isinstance(t, dict)}
        next_id = 1
        while next_id in existing_ids:
            next_id += 1

        tours.append({
            "id": next_id,
            "date": tour_date,
            "name": tour_name,
            "stops": stops,
            "employee_ids": selected_employee_ids,
            # Tour-Datenmodell ergänzt nur neue Keys und bleibt für Bestandsdaten kompatibel.
            "vehicle_id": selected_vehicle_id,
            "trailer_id": selected_trailer_id,
            "start_time": self.current_route_start_time,
            "route_mode": self.current_route_route_mode,
            "travel_time_cache": dict(self.current_route_travel_time_cache),
        })

        self._save_tours(tours)
        self.current_tour_id = next_id
        self.set_current_route_employee_ids(selected_employee_ids)
        self.set_current_route_resources(selected_vehicle_id, selected_trailer_id)
        self.show_toast("Tour gespeichert")

        self._refresh_all_markers()
        if self.current_selected_marker:
            self.on_marker_click(self.current_selected_marker)

        return True

    def update_tour_record(self, tour_id, tour_date: str, tour_name: str, employee_ids, start_time: str, vehicle_id: str, trailer_id=None):
        selected_employee_ids = []
        for value in employee_ids or []:
            key = str(value).strip()
            if key and key not in selected_employee_ids:
                selected_employee_ids.append(key)

        if not (1 <= len(selected_employee_ids) <= 2):
            raise ValueError("Bitte 1 oder 2 Mitarbeiter auswählen.")

        tour_date = _normalize_date_string(tour_date)
        if not tour_date:
            raise ValueError("Bitte ein Datum im Format DD.MM.YYYY eintragen.")

        parsed_start_time = parse_time(start_time)
        if parsed_start_time is None:
            raise ValueError("Bitte eine gültige Startzeit im Format HH:MM eintragen.")
        normalized_start_time = format_time(parsed_start_time)
        selected_vehicle_id = str(vehicle_id or "").strip()
        selected_trailer_id = str(trailer_id or "").strip() or None
        self._validate_tour_vehicle_selection(selected_vehicle_id, selected_trailer_id)

        parsed_tour_date = parse_date(tour_date)
        if parsed_tour_date and parsed_tour_date < datetime.now().date():
            messagebox.showwarning("Tour bearbeiten", "Das gewählte Datum liegt in der Vergangenheit.")

        conflicts = self._find_tour_resource_conflicts(
            tour_date,
            selected_vehicle_id,
            selected_trailer_id,
            exclude_tour_id=tour_id,
        )
        if not self._confirm_tour_resource_conflicts(conflicts):
            raise ValueError("Speichern wegen Ressourcenkonflikt abgebrochen.")

        tours = self._load_tours()
        updated = None
        for tour in tours:
            if tour.get("id") == tour_id:
                tour["date"] = tour_date
                tour["name"] = (tour_name or "").strip()
                tour["employee_ids"] = selected_employee_ids
                tour["vehicle_id"] = selected_vehicle_id
                tour["trailer_id"] = selected_trailer_id
                tour["start_time"] = normalized_start_time
                if self.current_tour_id == tour_id:
                    tour["travel_time_cache"] = dict(self.current_route_travel_time_cache)
                else:
                    tour["travel_time_cache"] = dict(tour.get("travel_time_cache") or {})
                updated = tour
                break

        if updated is None:
            raise ValueError("Die ausgewählte Tour wurde nicht gefunden.")

        self._save_tours(tours)

        if self.current_tour_id == tour_id:
            self.set_current_route_employee_ids(selected_employee_ids)
            self.set_current_route_resources(selected_vehicle_id, selected_trailer_id)
            self.current_route_start_time = normalized_start_time
            if hasattr(self, "lbl_current_tour"):
                self.lbl_current_tour.configure(text=self._tour_display_text(updated))
            try:
                _set_text_input_value(self.route_start_time_entry, normalized_start_time)
            except Exception:
                pass
            self._trigger_route_metrics_recalc(force_routing=False)

    def delete_tour_record(self, tour_id, *, confirm: bool = True) -> bool:
        if tour_id is None:
            return False

        tours = self._load_tours()
        target = next((tour for tour in tours if tour.get("id") == tour_id), None)
        if target is None:
            messagebox.showwarning("Tour löschen", "Die ausgewählte Tour wurde nicht gefunden.")
            return False

        tour_name = (target.get("name") or "").strip()
        tour_date = _normalize_date_string(target.get("date"))
        label = f"{tour_date}" + (f" – {tour_name}" if tour_name else "")
        prompt = f"Die Tour '{label}' wirklich löschen?" if label.strip(" –") else "Diese Tour wirklich löschen?"
        if confirm and not messagebox.askyesno("Tour löschen", prompt):
            return False

        remaining_tours = [tour for tour in tours if tour.get("id") != tour_id]
        self._save_tours(remaining_tours)

        if self.current_tour_id == tour_id:
            self.current_tour_id = None
            self.set_current_route_employee_ids([])
            if hasattr(self, "lbl_current_tour"):
                self.lbl_current_tour.configure(text="(manuell)")

        self._refresh_all_markers()
        if self.current_selected_marker:
            self.on_marker_click(self.current_selected_marker)
        return True

    def open_edit_current_tour(self):
        tour_id = getattr(self, "current_tour_id", None)
        if tour_id is None:
            messagebox.showwarning("Tour bearbeiten", "Bitte zuerst eine gespeicherte Tour laden oder auswählen.")
            return

        tour = None
        for item in self._load_tours():
            if item.get("id") == tour_id:
                tour = item
                break

        if tour is None:
            messagebox.showwarning("Tour bearbeiten", "Die aktuell geladene Tour wurde nicht gefunden.")
            return

        self.open_edit_tour_dialog(tour)

    def open_edit_tour_dialog(self, tour: dict):
        if not isinstance(tour, dict):
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title("Tour bearbeiten")
        dlg.geometry("640x700")
        dlg.resizable(True, True)
        dlg.configure(fg_color=Theme.BG)
        dlg.attributes("-topmost", True)

        shell = ctk.CTkFrame(dlg, corner_radius=18, fg_color=Theme.PANEL, border_width=1, border_color=Theme.BORDER)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        shell.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(shell, text="Liefertour bearbeiten", font=_font(16, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=(16, 10), sticky="w"
        )

        date_entry = ctk.CTkEntry(shell, height=36, corner_radius=12, placeholder_text="Datum (DD.MM.YYYY)")
        date_entry.grid(row=1, column=0, padx=16, pady=(0, 10), sticky="ew")
        date_entry.insert(0, _normalize_date_string(tour.get("date")))

        name_entry = ctk.CTkEntry(shell, height=36, corner_radius=12, placeholder_text="Name")
        name_entry.grid(row=2, column=0, padx=16, pady=(0, 10), sticky="ew")
        name_entry.insert(0, tour.get("name", ""))

        start_time_entry = TimeInput(shell, height=36)
        start_time_entry.grid(row=3, column=0, padx=16, pady=(0, 10), sticky="ew")
        start_time_entry.set(str(tour.get("start_time") or "08:00"))

        local_employee_ids = list(tour.get("employee_ids", []) or [])
        vehicle_labels, vehicle_options = self.build_vehicle_option_map(tour.get("vehicle_id"))
        trailer_labels, trailer_options = self.build_trailer_option_map(tour.get("trailer_id"))
        vehicle_var = tk.StringVar(value=self.get_vehicle_option_label(tour.get("vehicle_id"), vehicle_options))
        trailer_var = tk.StringVar(value=self.get_trailer_option_label(tour.get("trailer_id"), trailer_options))

        ctk.CTkLabel(shell, text="Fahrzeug *", font=_font(12, "bold"), text_color=Theme.TEXT).grid(
            row=4, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        vehicle_menu = ctk.CTkOptionMenu(shell, height=36, corner_radius=12, values=vehicle_labels, variable=vehicle_var)
        vehicle_menu.grid(row=5, column=0, padx=16, pady=(0, 10), sticky="ew")

        ctk.CTkLabel(shell, text="Anhänger", font=_font(12, "bold"), text_color=Theme.TEXT).grid(
            row=6, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        trailer_menu = ctk.CTkOptionMenu(shell, height=36, corner_radius=12, values=trailer_labels, variable=trailer_var)
        trailer_menu.grid(row=7, column=0, padx=16, pady=(0, 10), sticky="ew")

        resource_hint = ctk.CTkLabel(shell, text="", font=_font(12), text_color=Theme.WARNING, justify="left", wraplength=460)
        resource_hint.grid(row=8, column=0, padx=16, pady=(0, 8), sticky="w")

        def _selected_vehicle_id():
            return vehicle_options.get(vehicle_var.get())

        def _selected_trailer_id():
            return trailer_options.get(trailer_var.get())

        def _update_resource_hint(*_args):
            hint = self._resource_conflict_hint(
                date_entry.get().strip(),
                _selected_vehicle_id(),
                _selected_trailer_id(),
                exclude_tour_id=tour.get("id"),
            )
            resource_hint.configure(text=hint)

        employee_info = ctk.CTkLabel(
            shell,
            text=self.format_employee_summary(local_employee_ids),
            font=_font(12),
            text_color=Theme.SUBTEXT,
            justify="left",
        )
        employee_info.grid(row=9, column=0, padx=16, pady=(0, 8), sticky="w")

        def _set_dialog_employee_ids(ids):
            local_employee_ids.clear()
            local_employee_ids.extend(ids[:2])
            employee_info.configure(text=self.format_employee_summary(local_employee_ids))

        ctk.CTkButton(
            shell,
            text="Mitarbeiter wählen",
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: self.open_employee_picker(
                selected_ids=local_employee_ids,
                on_apply=_set_dialog_employee_ids,
                title="Mitarbeiter für Tour wählen",
            ),
        ).grid(row=10, column=0, padx=16, pady=(0, 10), sticky="ew")

        btns = ctk.CTkFrame(shell, fg_color="transparent")
        btns.grid(row=11, column=0, padx=16, pady=(6, 16), sticky="ew")
        btns.grid_columnconfigure((0, 1, 2), weight=1)

        date_entry.bind("<KeyRelease>", _update_resource_hint)
        vehicle_var.trace_add("write", _update_resource_hint)
        trailer_var.trace_add("write", _update_resource_hint)
        _update_resource_hint()

        def _save():
            try:
                self.update_tour_record(
                    tour.get("id"),
                    date_entry.get().strip(),
                    name_entry.get().strip(),
                    local_employee_ids,
                    start_time_entry.get().strip(),
                    _selected_vehicle_id(),
                    _selected_trailer_id(),
                )
            except ValueError as exc:
                messagebox.showwarning("Tour bearbeiten", str(exc))
                return
            except Exception as exc:
                messagebox.showerror("Tour bearbeiten", f"Tour konnte nicht gespeichert werden:\n{exc}")
                return
            try:
                dlg.destroy()
            except Exception:
                pass

        def _delete():
            if not self.delete_tour_record(tour.get("id"), confirm=True):
                return
            try:
                dlg.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            btns,
            text="Abbrechen",
            height=40,
            corner_radius=14,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=dlg.destroy,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            btns,
            text="Tour löschen",
            height=40,
            corner_radius=14,
            font=_font(13, "bold"),
            fg_color=Theme.DANGER,
            hover_color=Theme.DANGER_HOVER,
            command=_delete,
        ).grid(row=0, column=1, padx=8, sticky="ew")

        ctk.CTkButton(
            btns,
            text="Speichern",
            height=40,
            corner_radius=14,
            font=_font(13, "bold"),
            fg_color=Theme.SUCCESS,
            hover_color=Theme.SUCCESS_HOVER,
            command=_save,
        ).grid(row=0, column=2, padx=(8, 0), sticky="ew")

        dlg.grab_set()
        dlg.focus_force()

    def unload_current_tour(self):
        if not self.route_markers:
            return

        if self.route_path:
            try:
                self.route_path.delete()
            except Exception:
                pass
            self.route_path = None

        self.route_markers = []
        self.route_points = []
        self.current_route_stop_data = []
        self.current_route_segments = []
        self.current_route_travel_time_cache = {}
        self.current_route_distance_cache = {}
        self.current_tour_id = None
        self.set_current_route_employee_ids([])
        self.set_current_route_resources(None, None)
        self.current_route_start_time = "08:00"
        try:
            _set_text_input_value(self.route_start_time_entry, self.current_route_start_time)
        except Exception:
            pass

        if hasattr(self, "lbl_current_tour"):
            self.lbl_current_tour.configure(text="(manuell)")

        self._update_route_metrics_ui("Tour verlassen")
        self._refresh_all_markers()
        self.show_toast("Tour verlassen")

    def open_create_tour_dialog(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Neue Tour erstellen")
        dlg.geometry("640x760")
        dlg.resizable(True, True)
        dlg.configure(fg_color=Theme.BG)
        dlg.attributes("-topmost", True)

        shell = ctk.CTkFrame(dlg, corner_radius=18, fg_color=Theme.PANEL, border_width=1, border_color=Theme.BORDER)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        shell.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(shell, text="Neue Liefertour", font=_font(16, "bold"), text_color=Theme.TEXT).grid(
            row=0, column=0, padx=16, pady=(16, 10), sticky="w"
        )

        date_entry = ctk.CTkEntry(shell, placeholder_text="Datum (DD.MM.YYYY)", height=36, corner_radius=12)
        date_entry.grid(row=1, column=0, padx=16, pady=(0, 10), sticky="ew")

        name_entry = ctk.CTkEntry(shell, placeholder_text="Name (optional)", height=36, corner_radius=12)
        name_entry.grid(row=2, column=0, padx=16, pady=(0, 10), sticky="ew")

        start_time_entry = TimeInput(shell, height=36)
        start_time_entry.grid(row=3, column=0, padx=16, pady=(0, 10), sticky="ew")
        start_time_entry.set(self.current_route_start_time or "08:00")

        local_employee_ids = list(self.current_route_employee_ids)
        vehicle_labels, vehicle_options = self.build_vehicle_option_map(self.current_route_vehicle_id)
        trailer_labels, trailer_options = self.build_trailer_option_map(self.current_route_trailer_id)
        vehicle_var = tk.StringVar(value=self.get_vehicle_option_label(self.current_route_vehicle_id, vehicle_options))
        trailer_var = tk.StringVar(value=self.get_trailer_option_label(self.current_route_trailer_id, trailer_options))

        ctk.CTkLabel(shell, text="Fahrzeug *", font=_font(12, "bold"), text_color=Theme.TEXT).grid(
            row=4, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        vehicle_menu = ctk.CTkOptionMenu(shell, height=36, corner_radius=12, values=vehicle_labels, variable=vehicle_var)
        vehicle_menu.grid(row=5, column=0, padx=16, pady=(0, 10), sticky="ew")

        ctk.CTkLabel(shell, text="Anhänger", font=_font(12, "bold"), text_color=Theme.TEXT).grid(
            row=6, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        trailer_menu = ctk.CTkOptionMenu(shell, height=36, corner_radius=12, values=trailer_labels, variable=trailer_var)
        trailer_menu.grid(row=7, column=0, padx=16, pady=(0, 10), sticky="ew")

        resource_hint = ctk.CTkLabel(shell, text="", font=_font(12), text_color=Theme.WARNING, justify="left", wraplength=460)
        resource_hint.grid(row=8, column=0, padx=16, pady=(0, 8), sticky="w")

        def _selected_vehicle_id():
            return vehicle_options.get(vehicle_var.get())

        def _selected_trailer_id():
            return trailer_options.get(trailer_var.get())

        def _update_resource_hint(*_args):
            hint = self._resource_conflict_hint(
                date_entry.get().strip(),
                _selected_vehicle_id(),
                _selected_trailer_id(),
            )
            resource_hint.configure(text=hint)

        employee_info = ctk.CTkLabel(
            shell,
            text=self.format_employee_summary(local_employee_ids),
            font=_font(12),
            text_color=Theme.SUBTEXT,
            justify="left",
        )
        employee_info.grid(row=9, column=0, padx=16, pady=(0, 8), sticky="w")

        def _set_dialog_employee_ids(ids):
            local_employee_ids.clear()
            local_employee_ids.extend(ids[:2])
            employee_info.configure(text=self.format_employee_summary(local_employee_ids))

        ctk.CTkButton(
            shell,
            text="Mitarbeiter wählen",
            height=36,
            corner_radius=12,
            fg_color=Theme.PANEL_2,
            hover_color=Theme.BORDER,
            text_color=Theme.TEXT,
            command=lambda: self.open_employee_picker(
                selected_ids=local_employee_ids,
                on_apply=_set_dialog_employee_ids,
                title="Mitarbeiter für Tour wählen",
            ),
        ).grid(row=10, column=0, padx=16, pady=(0, 10), sticky="ew")

        btns = ctk.CTkFrame(shell, fg_color="transparent")
        btns.grid(row=11, column=0, padx=16, pady=(6, 16), sticky="ew")
        btns.grid_columnconfigure((0, 1), weight=1)

        date_entry.bind("<KeyRelease>", _update_resource_hint)
        vehicle_var.trace_add("write", _update_resource_hint)
        trailer_var.trace_add("write", _update_resource_hint)
        _update_resource_hint()

        def _create():
            tour_date = date_entry.get().strip()
            tour_name = name_entry.get().strip()
            parsed_start_time = parse_time(start_time_entry.get().strip())
            if parsed_start_time is None:
                messagebox.showwarning("Tour speichern", "Bitte eine gültige Startzeit im Format HH:MM eingeben.")
                return
            self.current_route_start_time = format_time(parsed_start_time)
            try:
                _set_text_input_value(self.route_start_time_entry, self.current_route_start_time)
            except Exception:
                pass
            before = len(self._load_tours())
            result = self.save_current_tour(
                tour_date,
                tour_name,
                employee_ids=local_employee_ids,
                vehicle_id=_selected_vehicle_id(),
                trailer_id=_selected_trailer_id(),
            )
            after = len(self._load_tours())
            if result is False or after <= before:
                return
            try:
                dlg.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            btns, text="Abbrechen", height=40, corner_radius=14,
            fg_color=Theme.PANEL_2, hover_color=Theme.BORDER, text_color=Theme.TEXT,
            command=dlg.destroy
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            btns, text="Erstellen", height=40, corner_radius=14, font=_font(13, "bold"),
            fg_color=Theme.SUCCESS, hover_color=Theme.SUCCESS_HOVER,
            command=_create
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        dlg.grab_set()
        dlg.focus_force()

    def apply_tour(self, tour: dict):
        if not tour or not isinstance(tour, dict):
            return

        self.current_tour_id = tour.get("id")
        self.set_current_route_employee_ids(tour.get("employee_ids", []))
        self.set_current_route_resources(tour.get("vehicle_id"), tour.get("trailer_id"))
        self.current_route_start_time = str(tour.get("start_time") or "08:00")
        self.current_route_route_mode = str(tour.get("route_mode") or "car")
        self.current_route_travel_time_cache = dict(tour.get("travel_time_cache") or {})
        self.current_route_distance_cache = {}
        self.current_route_stop_data = [dict(stop) for stop in (tour.get("stops", []) or [])]
        try:
            _set_text_input_value(self.route_start_time_entry, self.current_route_start_time)
        except Exception:
            pass

        stops = tour.get("stops", [])
        if not stops or len(stops) < 1:
            messagebox.showwarning("Tour", "Diese Tour enthält zu wenige Stopps.")
            return

        if self.route_path:
            self.route_path.delete()
            self.route_path = None
        self.route_points = []
        self.route_markers = []

        missing = 0

        self._ensure_depot_markers_exist()
        if self.depot_start_marker and self.depot_end_marker:
            self.route_markers = [self.depot_start_marker]

        for s in stops:
            marker = None
            if isinstance(s, dict):
                marker = self._find_marker_for_stop(s)

            if marker is None:
                missing += 1
                continue

            self.route_markers.append(marker)

        if self.depot_end_marker:
            self.route_markers.append(self.depot_end_marker)

        self._rebuild_route_from_markers()
        self._sync_current_route_stop_data_from_markers(preferred_stops=stops)
        self._trigger_route_metrics_recalc(force_routing=False)

        if len(self.route_points) < 2:
            messagebox.showwarning("Tour", "Tour konnte nicht angezeigt werden (Stopps fehlen).")
            return

        self.show_page("map")

        try:
            self.map_widget.set_position(self.route_points[0][0], self.route_points[0][1])
            self.map_widget.set_zoom(11)
        except Exception:
            pass

        if missing:
            messagebox.showwarning("Tour", f"{missing} Stopps konnten nicht zugeordnet werden (Pins fehlen).")

        self._refresh_all_markers()
        self.show_toast("Tour geladen")

    def show_tour_of_selected_pin(self):
        marker = getattr(self, "current_selected_marker", None)
        if not marker:
            messagebox.showwarning("Tour", "Kein Pin ausgewählt.")
            return

        tour = getattr(self, "_selected_pin_tour", None)
        if tour is None:
            tour = self._pin_used_in_any_tour(marker)

        if not tour:
            messagebox.showwarning("Tour", "Dieser Pin ist aktuell in keiner Tour eingeplant.")
            return

        self.apply_tour(tour)

    def remove_selected_pin_from_tour(self):
        marker = getattr(self, "current_selected_marker", None)
        if not marker:
            messagebox.showwarning("Tour", "Kein Pin ausgewählt.")
            return

        # Remove from persisted tour (not just current route)
        key = self._marker_key(marker)
        tour = getattr(self, "_selected_pin_tour", None) or self._pin_used_in_any_tour(marker)
        if not tour:
            messagebox.showwarning("Tour", "Dieser Pin ist aktuell in keiner Tour eingeplant.")
            return

        tours = self._load_tours()
        updated = False
        for t in tours:
            if t.get("id") == tour.get("id"):
                new_stops = []
                for s in t.get("stops", []):
                    if not isinstance(s, dict):
                        continue
                    s_key = None
                    if (s.get("auftragsnummer") or "").strip():
                        s_key = ("auftrag", str(s.get("auftragsnummer")).strip())
                    elif s.get("lat") is not None and s.get("lng") is not None:
                        try:
                            s_key = ("coord", round(float(s["lat"]), 7), round(float(s["lng"]), 7))
                        except Exception:
                            s_key = None

                    if s_key != key:
                        new_stops.append(s)
                    else:
                        updated = True

                t["stops"] = new_stops
                break

        if not updated:
            messagebox.showinfo("Tour", "Dieser Pin wurde in der Tour nicht gefunden.")
            return

        self._save_tours(tours)
        self.show_toast("Pin aus Tour entfernt")

        # refresh marker colors + UI
        self._refresh_all_markers()
        if self.current_selected_marker:
            self.on_marker_click(self.current_selected_marker)

        # if that tour is currently applied, re-apply (optional but keeps route clean)
        try:
            cur_label = getattr(self, "lbl_current_tour", None)
            if cur_label is not None and cur_label.cget("text") == self._tour_display_text(tour):
                self.apply_tour(tour)  # may now have fewer stops
        except Exception:
            pass

    # ---------- Marker refresh ----------
    def _refresh_all_markers(self):
        existing = list(self.marker_list)
        route_keys = [self._marker_key(m) for m in getattr(self, "route_markers", [])]
        route_key_set = set(route_keys)
        selected_key = self._marker_key(self.current_selected_marker) if self.current_selected_marker else None
        tour_usage = self._build_tour_usage_cache()
        selected_tour = tour_usage.get(selected_key) if selected_key else None
        selected_statuses = set(getattr(self, "map_filter_selected_statuses", set()))
        selected_delivery_types = {
            _normalize_delivery_type(value)
            for value in getattr(self, "map_filter_selected_delivery_types", set())
        }
        only_current_tour = bool(getattr(self, "map_filter_only_current_tour", False))
        selected_hidden = False

        self.marker_list.clear()
        new_by_key = {}

        for m in existing:
            try:
                lat, lng = m.position
                data = getattr(m, "data", {}) or {}
                status = getattr(m, "status", "nicht festgelegt")
                delivery_type = _normalize_delivery_type(data.get("Lieferart"))
                data["Lieferart"] = delivery_type

                is_system = bool(getattr(m, "is_system", False))
                anchor_id = getattr(m, "route_anchor_id", None)

                key_old = self._marker_key(m)
                in_tour = key_old in tour_usage
                if (not is_system) and _is_non_map_delivery_type(delivery_type):
                    try:
                        m.delete()
                    except Exception:
                        pass
                    if selected_key is not None and key_old == selected_key:
                        selected_hidden = True
                    continue
                marker_filter_statuses = {status}
                if in_tour:
                    marker_filter_statuses.add("Bereits eingeplant")
                if only_current_tour:
                    show_marker = bool(is_system or key_old in route_key_set)
                else:
                    matches_status = bool(marker_filter_statuses.intersection(selected_statuses))
                    matches_delivery_type = delivery_type in selected_delivery_types
                    show_marker = bool(
                        is_system
                        or key_old in route_key_set
                        or (matches_status and matches_delivery_type)
                    )
                if not show_marker:
                    try:
                        m.delete()
                    except Exception:
                        pass
                    self.marker_list.append(m)
                    if selected_key is not None and key_old == selected_key:
                        selected_hidden = True
                    continue

                m.delete()

                if is_system and anchor_id in ("depot_start", "depot_end"):
                    icon = self._get_depot_marker_icon(anchor_id)
                else:
                    icon = self._get_marker_icon(status, in_tour, delivery_type=delivery_type)

                new_m = self.map_widget.set_marker(lat, lng, text="", icon=icon, command=self.on_marker_click)

                new_m.status = status
                new_m.data = data
                new_m.full_info = self._build_popup_text(data)
                new_m.email = data.get("Email", "")
                new_m.auftragsnummer = data.get("Auftragsnummer", "")
                new_m.import_id = str(data.get("ImportID") or "").strip()

                if is_system:
                    new_m.is_system = True
                    new_m.route_anchor_id = anchor_id
                    if anchor_id == "depot_start":
                        self.depot_start_marker = new_m
                    elif anchor_id == "depot_end":
                        self.depot_end_marker = new_m

                self._style_marker_label(new_m)
                self.marker_list.append(new_m)
                new_by_key[self._marker_key(new_m)] = new_m
            except Exception:
                pass

        self.route_markers = [new_by_key[k] for k in route_keys if k in new_by_key]

        if selected_key and selected_key in new_by_key:
            self.current_selected_marker = new_by_key[selected_key]
            self._selected_pin_tour = selected_tour
        else:
            self.current_selected_marker = None
            self._selected_pin_tour = None

        self._rebuild_route_from_markers()
        self._trigger_route_metrics_recalc(force_routing=False)

        if self.current_selected_marker:
            try:
                self.on_marker_click(self.current_selected_marker)
            except Exception:
                pass
        elif selected_hidden:
            self.hide_info_card()

    # ---------- Zoom watcher ----------
    def start_zoom_watch(self):
        if self._zoom_watch_running:
            return
        self._zoom_watch_running = True
        self._check_zoom_loop()

    def _get_current_zoom(self):
        try:
            if hasattr(self.map_widget, "zoom"):
                return int(self.map_widget.zoom)
        except Exception:
            pass
        try:
            if hasattr(self.map_widget, "get_zoom"):
                return int(self.map_widget.get_zoom())
        except Exception:
            pass
        return None

    def _check_zoom_loop(self):
        next_delay_ms = 500
        try:
            if getattr(self, "current_page", "") != "map":
                next_delay_ms = 700
                return
            if not hasattr(self, "map_widget") or self.map_widget is None:
                return

            zoom = self._get_current_zoom()
            if zoom is None:
                return

            if zoom != self._last_zoom:
                self._last_zoom = zoom
                self._schedule_marker_size_update(zoom)
        finally:
            self.after(next_delay_ms, self._check_zoom_loop)

    def _schedule_marker_size_update(self, zoom: int, delay_ms: int = 260):
        if self._pending_marker_size_job is not None:
            try:
                self.after_cancel(self._pending_marker_size_job)
            except Exception:
                pass
        self._pending_marker_size_job = self.after(
            delay_ms,
            lambda z=int(zoom): self._apply_marker_size_for_zoom(z),
        )

    def _apply_marker_size_for_zoom(self, zoom: int):
        self._pending_marker_size_job = None
        if getattr(self, "current_page", "") != "map":
            return
        if self._is_window_resizing:
            self._schedule_marker_size_update(zoom, delay_ms=240)
            return

        if zoom <= 9:
            size = 18
        elif zoom <= 12:
            size = 20
        elif zoom <= 15:
            size = 28
        else:
            size = 28

        if size == self.marker_icon_size:
            return

        self.marker_icon_size = size
        self.marker_icons = self._build_marker_icons(size)
        self._refresh_all_markers()


if __name__ == "__main__":
    if _run_auxiliary_mode_from_argv():
        raise SystemExit(0)
    app = ModernApp()
    app.mainloop()
