from __future__ import annotations

import logging
import threading
from tkinter import messagebox

import customtkinter as ctk

from config.update_config import APPINSTALLER_URL, SUPPORT_URL
from services.version_service import (
    check_update_source_reachable,
    format_auto_update_settings,
    get_auto_update_settings,
    get_runtime_update_context,
    is_auto_update_settings_supported,
    open_support_url,
    set_auto_update_check_on_launch,
    trigger_update_installation,
)

logger = logging.getLogger(__name__)


class UpdatePage(ctk.CTkFrame):
    # Testfälle:
    # - MSIX installiert via .appinstaller -> Button öffnet App Installer und Updates werden gefunden.
    # - Nicht-MSIX (portable) -> UI zeigt Hinweis, Button öffnet Browser.
    # - ms-appinstaller Protokoll deaktiviert -> Browser-Fallback + Hinweis.
    # - Kein Internet -> klare Meldung, kein Freeze.

    def __init__(self, master, app, theme, font_factory):
        super().__init__(master, fg_color=theme.BG)
        self.app = app
        self.theme = theme
        self.font = font_factory
        self._refresh_running = False
        self._toggle_update_in_progress = False
        self._state: dict = {}

        self.internet_var = ctk.StringVar(value="Prüfung läuft ...")
        self.installation_var = ctk.StringVar(value="Wird ermittelt ...")
        self.version_var = ctk.StringVar(value="Wird ermittelt ...")
        self.url_var = ctk.StringVar(value=APPINSTALLER_URL)
        self.status_var = ctk.StringVar(value="Update-Status wird geladen ...")
        self.auto_update_info_var = ctk.StringVar(value="AutoUpdate-Einstellungen werden geprüft ...")
        self.check_on_launch_var = ctk.BooleanVar(value=False)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        topbar = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=self.theme.PANEL,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        topbar.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 12))
        topbar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(topbar, text="Updates", font=self.font(18, "bold"), text_color=self.theme.TEXT).grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )

        ctk.CTkButton(
            topbar,
            text="Status aktualisieren",
            width=150,
            height=36,
            corner_radius=12,
            fg_color=self.theme.PANEL_2,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=self.refresh,
        ).grid(row=0, column=2, padx=(8, 14), pady=10)

        shell = ctk.CTkScrollableFrame(
            self,
            corner_radius=18,
            fg_color=self.theme.PANEL,
            border_width=1,
            border_color=self.theme.BORDER,
            scrollbar_fg_color=self.theme.SCROLLBAR_TRACK,
            scrollbar_button_color=self.theme.SCROLLBAR_BUTTON,
            scrollbar_button_hover_color=self.theme.SCROLLBAR_HOVER,
        )
        shell.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        shell.grid_columnconfigure(0, weight=1)

        status_card = ctk.CTkFrame(
            shell,
            corner_radius=16,
            fg_color=self.theme.PANEL_2,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        status_card.grid(row=0, column=0, padx=14, pady=(14, 10), sticky="ew")
        status_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            status_card,
            text="Update-Status",
            font=self.font(16, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=0, column=0, columnspan=2, padx=16, pady=(16, 10), sticky="w")

        self._status_row(status_card, 1, "Internet", self.internet_var)
        self._status_row(status_card, 2, "Installationsart", self.installation_var)
        self._status_row(status_card, 3, "Version", self.version_var)

        ctk.CTkLabel(
            status_card,
            text="AppInstaller URL",
            font=self.font(12, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=4, column=0, padx=16, pady=(8, 6), sticky="w")

        url_row = ctk.CTkFrame(status_card, fg_color="transparent")
        url_row.grid(row=4, column=1, padx=(0, 16), pady=(8, 6), sticky="ew")
        url_row.grid_columnconfigure(0, weight=1)

        self.url_entry = ctk.CTkEntry(
            url_row,
            textvariable=self.url_var,
            height=38,
            corner_radius=12,
            fg_color=self.theme.PANEL,
            border_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
        )
        self.url_entry.grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            url_row,
            text="Kopieren",
            width=100,
            height=38,
            corner_radius=12,
            fg_color=self.theme.PANEL,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=self._copy_appinstaller_url,
        ).grid(row=0, column=1, sticky="ew")

        self.status_label = ctk.CTkLabel(
            status_card,
            textvariable=self.status_var,
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
            justify="left",
            wraplength=960,
        )
        self.status_label.grid(row=5, column=0, columnspan=2, padx=16, pady=(10, 16), sticky="w")

        actions_card = ctk.CTkFrame(
            shell,
            corner_radius=16,
            fg_color=self.theme.PANEL_2,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        actions_card.grid(row=1, column=0, padx=14, pady=(0, 10), sticky="ew")
        actions_card.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(
            actions_card,
            text="Aktionen",
            font=self.font(16, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=0, column=0, columnspan=3, padx=16, pady=(16, 10), sticky="w")

        self.update_now_button = ctk.CTkButton(
            actions_card,
            text="Jetzt nach Updates suchen",
            height=40,
            corner_radius=12,
            fg_color=self.theme.ACCENT,
            hover_color=self.theme.ACCENT_HOVER,
            text_color=("white", "white"),
            command=self._start_update_trigger,
        )
        self.update_now_button.grid(row=1, column=0, padx=(16, 8), pady=(0, 12), sticky="ew")

        self.show_settings_button = ctk.CTkButton(
            actions_card,
            text="Update-Einstellungen anzeigen",
            height=40,
            corner_radius=12,
            fg_color=self.theme.PANEL,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=self._show_auto_update_settings_dialog,
        )
        self.show_settings_button.grid(row=1, column=1, padx=8, pady=(0, 12), sticky="ew")

        self.help_button = ctk.CTkButton(
            actions_card,
            text="Hilfe",
            height=40,
            corner_radius=12,
            fg_color=self.theme.PANEL,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=self._open_help,
        )
        self.help_button.grid(row=1, column=2, padx=(8, 16), pady=(0, 12), sticky="ew")

        auto_card = ctk.CTkFrame(
            shell,
            corner_radius=16,
            fg_color=self.theme.PANEL_2,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        auto_card.grid(row=2, column=0, padx=14, pady=(0, 14), sticky="ew")
        auto_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            auto_card,
            text="Automatische Updates",
            font=self.font(16, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=0, column=0, padx=16, pady=(16, 6), sticky="w")

        self.auto_update_switch = ctk.CTkSwitch(
            auto_card,
            text="Beim Start nach Updates suchen",
            variable=self.check_on_launch_var,
            onvalue=True,
            offvalue=False,
            text_color=self.theme.TEXT,
            command=self._toggle_check_on_launch,
        )
        self.auto_update_switch.grid(row=1, column=0, padx=16, pady=(0, 8), sticky="w")

        self.auto_update_info = ctk.CTkLabel(
            auto_card,
            textvariable=self.auto_update_info_var,
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
            justify="left",
            wraplength=960,
        )
        self.auto_update_info.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="w")

        self.refresh()

    def _status_row(self, master, row: int, label: str, variable):
        ctk.CTkLabel(
            master,
            text=label,
            font=self.font(12, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=row, column=0, padx=16, pady=6, sticky="w")
        ctk.CTkLabel(
            master,
            textvariable=variable,
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
            justify="left",
            wraplength=840,
        ).grid(row=row, column=1, padx=(0, 16), pady=6, sticky="w")

    def refresh(self):
        if self._refresh_running:
            return
        self._refresh_running = True
        self.status_var.set("Update-Status wird geladen ...")
        self.internet_var.set("Prüfung läuft ...")

        def _worker():
            try:
                context = get_runtime_update_context()
                internet = check_update_source_reachable()
                auto_supported = is_auto_update_settings_supported()
                auto_settings = None
                if context.get("is_msix") and auto_supported and context.get("package_family_name"):
                    auto_settings = get_auto_update_settings(context["package_family_name"], show_update_availability=True)

                payload = {
                    "context": context,
                    "internet": internet,
                    "auto_supported": auto_supported,
                    "auto_settings": auto_settings,
                }
                self.after(0, lambda: self._apply_refresh_result(payload))
            except Exception as exc:
                logger.exception("Update status refresh failed.")
                self.after(0, lambda: self._apply_refresh_error(str(exc) or "Unbekannter Fehler beim Laden des Update-Status."))

        threading.Thread(target=_worker, daemon=True).start()

    def on_show(self):
        self.refresh()

    def _apply_refresh_result(self, payload: dict):
        self._refresh_running = False
        self._state = payload

        context = payload.get("context", {}) or {}
        internet = payload.get("internet", {}) or {}
        auto_supported = bool(payload.get("auto_supported"))
        auto_settings = payload.get("auto_settings")

        internet_text = "OK" if internet.get("ok") else "Nicht verfügbar"
        internet_detail = str(internet.get("detail") or "").strip()
        self.internet_var.set(f"{internet_text}" + (f" ({internet_detail})" if internet_detail else ""))

        installation_type = str(context.get("installation_type") or "unknown")
        install_label = "MSIX" if installation_type == "msix" else "Nicht-MSIX"
        self.installation_var.set(install_label)
        self.version_var.set(str(context.get("version") or "Unbekannt"))
        self.url_var.set(str(context.get("appinstaller_url") or APPINSTALLER_URL))

        status_lines = [
            f"Installationsart: {install_label}",
            f"Version: {self.version_var.get()}",
        ]
        if installation_type != "msix":
            status_lines.append("Auto-Update ist nur verfügbar, wenn die App via .appinstaller (MSIX) installiert wurde.")
        else:
            status_lines.append("Updates werden automatisch über Windows App Installer geprüft, wenn die .appinstaller-Datei so konfiguriert ist.")
        if not internet.get("ok"):
            status_lines.append("Kein Update möglich, solange die AppInstaller-Quelle nicht erreichbar ist.")
        self.status_var.set("\n".join(status_lines))

        if installation_type == "msix" and auto_supported:
            self.show_settings_button.configure(state="normal")
            check_on_launch = bool(auto_settings.get("CheckOnLaunch")) if isinstance(auto_settings, dict) else False
            self.check_on_launch_var.set(check_on_launch)
            self.auto_update_switch.configure(state="normal")
            self.auto_update_info_var.set(format_auto_update_settings(auto_settings))
        elif installation_type == "msix":
            self.show_settings_button.configure(state="disabled")
            self.check_on_launch_var.set(False)
            self.auto_update_switch.configure(state="disabled")
            self.auto_update_info_var.set(
                "Die Windows-Cmdlets für AutoUpdate-Einstellungen sind auf diesem System nicht verfügbar."
            )
        else:
            self.show_settings_button.configure(state="disabled")
            self.check_on_launch_var.set(False)
            self.auto_update_switch.configure(state="disabled")
            self.auto_update_info_var.set(
                "Diese Installation ist nicht als MSIX erkannt. Für automatische Updates bitte über die .appinstaller-Datei installieren."
            )

    def _apply_refresh_error(self, detail: str):
        self._refresh_running = False
        self._state = {}
        self.internet_var.set("Fehler")
        self.installation_var.set("Fehler beim Laden")
        self.version_var.set("Fehler beim Laden")
        self.status_var.set(f"Update-Status konnte nicht geladen werden.\n{detail}")
        self.show_settings_button.configure(state="disabled")
        self.check_on_launch_var.set(False)
        self.auto_update_switch.configure(state="disabled")
        self.auto_update_info_var.set("Die Update-Informationen konnten nicht geladen werden.")

    def _copy_appinstaller_url(self):
        value = self.url_var.get().strip()
        if not value:
            messagebox.showwarning("Updates", "Keine AppInstaller-URL konfiguriert.")
            return
        self.clipboard_clear()
        self.clipboard_append(value)
        messagebox.showinfo("Updates", "Die AppInstaller-URL wurde in die Zwischenablage kopiert.")

    def _start_update_trigger(self):
        state = self._state or {}
        internet = state.get("internet", {}) or {}
        if not internet.get("ok"):
            messagebox.showwarning("Updates", "Kein Update möglich: Die AppInstaller-Quelle ist aktuell nicht erreichbar.")
            return

        self.update_now_button.configure(state="disabled")
        self.status_var.set("Update wird gestartet ...")
        context = state.get("context", {}) or {}
        prefer_appinstaller = bool(context.get("is_msix"))

        def _worker():
            result = trigger_update_installation(prefer_appinstaller=prefer_appinstaller)
            self.after(0, lambda: self._finish_update_trigger(result))

        threading.Thread(target=_worker, daemon=True).start()

    def _finish_update_trigger(self, result: dict):
        self.update_now_button.configure(state="normal")
        detail = str(result.get("detail") or "Unbekanntes Ergebnis.")
        self.status_var.set(detail)
        if result.get("ok"):
            messagebox.showinfo("Updates", detail)
        else:
            messagebox.showwarning("Updates", detail)

    def _show_auto_update_settings_dialog(self):
        context = (self._state or {}).get("context", {}) or {}
        if not context.get("is_msix"):
            messagebox.showinfo("Updates", "Diese Installation ist nicht als MSIX erkannt.")
            return
        settings = (self._state or {}).get("auto_settings")
        messagebox.showinfo("Update-Einstellungen", format_auto_update_settings(settings))

    def _open_help(self):
        if not SUPPORT_URL:
            messagebox.showinfo("Updates", "Keine Support-URL konfiguriert.")
            return
        if not open_support_url():
            messagebox.showwarning("Updates", "Die Hilfeseite konnte nicht geöffnet werden.")

    def _toggle_check_on_launch(self):
        if self._toggle_update_in_progress:
            return

        context = (self._state or {}).get("context", {}) or {}
        package_family_name = str(context.get("package_family_name") or "").strip()
        if not context.get("is_msix") or not package_family_name:
            self.check_on_launch_var.set(False)
            return

        desired_state = bool(self.check_on_launch_var.get())
        self._toggle_update_in_progress = True
        self.auto_update_switch.configure(state="disabled")
        self.auto_update_info_var.set("Windows Update-Einstellung wird gespeichert ...")

        def _worker():
            success, detail = set_auto_update_check_on_launch(package_family_name, desired_state)
            self.after(0, lambda: self._finish_toggle_check_on_launch(success, detail, desired_state))

        threading.Thread(target=_worker, daemon=True).start()

    def _finish_toggle_check_on_launch(self, success: bool, detail: str, desired_state: bool):
        self._toggle_update_in_progress = False
        if not success:
            self.check_on_launch_var.set(not desired_state)
            self.auto_update_info_var.set(detail)
            messagebox.showwarning("Updates", detail)
        else:
            self.check_on_launch_var.set(desired_state)
            self.auto_update_info_var.set(detail)
        self.refresh()
