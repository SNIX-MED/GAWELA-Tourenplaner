from __future__ import annotations
import threading
from datetime import datetime, timezone
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

import customtkinter as ctk

from backup_manager import BackupManager
from settings_manager import SettingsManager


class QuickAccessPicker(ctk.CTkFrame):
    def __init__(self, master, *, values, variable, command, theme, font_factory):
        super().__init__(master, fg_color="transparent")
        self.values = [str(value) for value in values]
        self.variable = variable
        self.command = command
        self.theme = theme
        self.font = font_factory
        self._popup = None
        self._outside_click_binding_id = None

        self.grid_columnconfigure(0, weight=1)

        self.button = ctk.CTkButton(
            self,
            text=self._button_text(),
            height=38,
            corner_radius=12,
            font=self.font(12, "bold"),
            fg_color=self.theme.PANEL,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            anchor="w",
            command=self.toggle_popup,
        )
        self.button.grid(row=0, column=0, sticky="ew")

    def _button_text(self) -> str:
        value = str(self.variable.get() or "").strip() or "Kein Eintrag"
        return f"{value} ▾"

    def refresh(self):
        self.button.configure(text=self._button_text())

    def set(self, value: str):
        self.variable.set(value)
        self.refresh()

    def toggle_popup(self):
        popup = self._popup
        if popup is not None:
            try:
                if popup.winfo_exists():
                    self.close_popup()
                    return
            except tk.TclError:
                pass
        self.open_popup()

    def open_popup(self):
        if self._popup is not None:
            return

        root = self.winfo_toplevel()
        popup = ctk.CTkToplevel(root)
        self._popup = popup
        popup.title("Schnellzugriff auswählen")
        popup.resizable(False, False)
        popup.attributes("-topmost", True)
        popup.configure(fg_color=self.theme.BG)

        self.update_idletasks()
        x = self.button.winfo_rootx()
        y = self.button.winfo_rooty() + self.button.winfo_height() + 8
        popup.geometry(f"320x280+{x}+{y}")

        popup.grid_columnconfigure(0, weight=1)
        popup.grid_rowconfigure(0, weight=1)
        popup.protocol("WM_DELETE_WINDOW", self.close_popup)

        shell = ctk.CTkFrame(
            popup,
            corner_radius=16,
            fg_color=self.theme.PANEL,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        shell.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(
            shell,
            text="Schnellzugriff auswählen",
            font=self.font(13, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=0, column=0, padx=14, pady=(12, 8), sticky="w")

        scroll = ctk.CTkScrollableFrame(
            shell,
            fg_color="transparent",
            scrollbar_fg_color=self.theme.SCROLLBAR_TRACK,
            scrollbar_button_color=self.theme.SCROLLBAR_BUTTON,
            scrollbar_button_hover_color=self.theme.SCROLLBAR_HOVER,
        )
        scroll.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        scroll.grid_columnconfigure(0, weight=1)

        current = str(self.variable.get() or "").strip()

        for index, value in enumerate(self.values):
            selected = value == current
            option = ctk.CTkButton(
                scroll,
                text=value,
                height=34,
                corner_radius=12,
                font=self.font(12, "bold" if selected else "normal"),
                fg_color=self.theme.SELECTION if selected else self.theme.PANEL_2,
                hover_color=self.theme.BORDER,
                text_color=self.theme.TEXT,
                anchor="w",
                command=lambda selected_value=value: self._select(selected_value),
            )
            option.grid(row=index, column=0, padx=6, pady=4 if index else (2, 4), sticky="ew")

        popup.after(10, lambda: popup.focus_force() if popup.winfo_exists() else None)
        self._outside_click_binding_id = root.bind("<Button-1>", self._handle_global_click, add="+")

    def _handle_global_click(self, event=None):
        popup = self._popup
        if popup is None:
            return
        widget = getattr(event, "widget", None)
        if widget is None:
            self.close_popup()
            return
        try:
            widget_toplevel = widget.winfo_toplevel()
        except Exception:
            self.close_popup()
            return
        if widget_toplevel == popup:
            return
        if widget == self.button:
            return
        try:
            if str(widget).startswith(str(self)):
                return
        except Exception:
            pass
        self.close_popup()

    def _select(self, value: str):
        self.set(value)
        if callable(self.command):
            self.command(value)
        self.close_popup()

    def close_popup(self):
        popup = self._popup
        self._popup = None
        root = self.winfo_toplevel()
        if self._outside_click_binding_id is not None:
            try:
                root.unbind("<Button-1>", self._outside_click_binding_id)
            except tk.TclError:
                pass
            self._outside_click_binding_id = None
        if popup is None:
            return
        try:
            if popup.winfo_exists():
                popup.destroy()
        except tk.TclError:
            pass


class SettingsPage(ctk.CTkFrame):
    RESTORE_GROUPS = [
        ("orders", "Aufträge & Adressen", "pins.json"),
        ("tours", "Liefertouren", "tours.json"),
        ("employees", "Mitarbeiter", "data/employees.json"),
        ("vehicles", "Fahrzeuge", "data/vehicles.json"),
        ("settings", "Einstellungen", "settings.json"),
        ("misc", "Zusatzdaten", "geocode_cache.json / config.json"),
        ("other_data", "Weitere Daten", "sonstige Dateien in data/"),
    ]

    def __init__(self, master, app, theme, font_factory):
        super().__init__(master, fg_color=theme.BG)
        self.app = app
        self.theme = theme
        self.font = font_factory
        self.settings_manager = SettingsManager(Path(getattr(self.app, "config_dir", getattr(self.app, "base_dir", "."))))
        self._backup_running = False
        self._appearance_buttons = {}
        self._xml_info = None
        self._theme_info = None
        self._quick_access_info = None
        self._quick_access_option_map = dict(self.app.get_quick_access_options()) if hasattr(self.app, "get_quick_access_options") else {}
        self._quick_access_labels = [label for _key, label in self.app.get_quick_access_options()] if hasattr(self.app, "get_quick_access_options") else ["Kein Eintrag"]
        self._quick_access_vars = [ctk.StringVar(value="Kein Eintrag") for _ in range(4)]
        self._quick_access_menus = []
        self._backup_enabled_var = tk.BooleanVar(value=False)
        self._auto_backup_enabled_var = tk.BooleanVar(value=False)
        self._backup_dir_var = ctk.StringVar(value="")
        self._backup_mode_var = ctk.StringVar(value="full")
        self._backup_retention_var = ctk.StringVar(value="30")
        self._auto_backup_interval_var = ctk.StringVar(value="7")
        self._backup_status_var = ctk.StringVar(value="Kein Backup ausgeführt.")

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

        ctk.CTkLabel(topbar, text="Einstellungen", font=self.font(18, "bold"), text_color=self.theme.TEXT).grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )

        ctk.CTkButton(
            topbar,
            text="Aktualisieren",
            width=120,
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
        self.shell = shell

        self._build_content()
        self.refresh()

    def _build_content(self):
        appearance_card = ctk.CTkFrame(
            self.shell,
            corner_radius=16,
            fg_color=self.theme.PANEL_2,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        appearance_card.grid(row=0, column=0, padx=14, pady=(14, 10), sticky="ew")
        appearance_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            appearance_card,
            text="Darstellung",
            font=self.font(16, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=0, column=0, padx=16, pady=(16, 6), sticky="w")

        self._theme_info = ctk.CTkLabel(
            appearance_card,
            text="",
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
            justify="left",
        )
        self._theme_info.grid(row=1, column=0, padx=16, pady=(0, 10), sticky="w")

        mode_row = ctk.CTkFrame(appearance_card, fg_color="transparent")
        mode_row.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
        mode_row.grid_columnconfigure((0, 1, 2), weight=1)

        for col, (mode, label) in enumerate((("System", "System"), ("Light", "Hell"), ("Dark", "Dunkel"))):
            button = ctk.CTkButton(
                mode_row,
                text=label,
                height=40,
                corner_radius=12,
                fg_color=self.theme.PANEL,
                hover_color=self.theme.BORDER,
                text_color=self.theme.TEXT,
                border_width=1,
                border_color=self.theme.BORDER,
                command=lambda value=mode: self._set_appearance(value),
            )
            button.grid(row=0, column=col, padx=(0, 8) if col < 2 else (8, 0), sticky="ew")
            self._appearance_buttons[mode] = button

        data_card = ctk.CTkFrame(
            self.shell,
            corner_radius=16,
            fg_color=self.theme.PANEL_2,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        data_card.grid(row=1, column=0, padx=14, pady=(0, 14), sticky="ew")
        data_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            data_card,
            text="Daten & Speicher",
            font=self.font(16, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=0, column=0, padx=16, pady=(16, 6), sticky="w")

        self._xml_info = ctk.CTkLabel(
            data_card,
            text="",
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
            justify="left",
            wraplength=900,
        )
        self._xml_info.grid(row=1, column=0, padx=16, pady=(0, 10), sticky="w")

        xml_row = ctk.CTkFrame(data_card, fg_color="transparent")
        xml_row.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
        xml_row.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            xml_row,
            text="XML-Ordner wählen",
            height=40,
            corner_radius=12,
            fg_color=self.theme.PANEL,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=self._select_xml_folder_and_refresh,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            xml_row,
            text="Ordner importieren",
            height=40,
            corner_radius=12,
            fg_color=self.theme.ACCENT,
            hover_color=self.theme.ACCENT_HOVER,
            text_color=("white", "white"),
            command=self.app.import_xml_from_folder,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        quick_card = ctk.CTkFrame(
            self.shell,
            corner_radius=16,
            fg_color=self.theme.PANEL_2,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        quick_card.grid(row=2, column=0, padx=14, pady=(0, 14), sticky="ew")
        quick_card.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(
            quick_card,
            text="Schnellzugriff",
            font=self.font(16, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=0, column=0, columnspan=2, padx=16, pady=(16, 6), sticky="w")

        self._quick_access_info = ctk.CTkLabel(
            quick_card,
            text="",
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
            justify="left",
            wraplength=900,
        )
        self._quick_access_info.grid(row=1, column=0, columnspan=2, padx=16, pady=(0, 10), sticky="w")

        for index, variable in enumerate(self._quick_access_vars, start=1):
            row = 2 + ((index - 1) // 2) * 2
            col = (index - 1) % 2
            padx = (16, 8) if col == 0 else (8, 16)

            ctk.CTkLabel(
                quick_card,
                text=f"Schnellzugriff {index}",
                font=self.font(12, "bold"),
                text_color=self.theme.TEXT,
            ).grid(row=row, column=col, padx=padx, pady=(0, 6), sticky="w")

            menu = QuickAccessPicker(
                quick_card,
                values=self._quick_access_labels,
                variable=variable,
                command=lambda _value, slot=index - 1: self._on_quick_access_change(slot),
                theme=self.theme,
                font_factory=self.font,
            )
            menu.grid(row=row + 1, column=col, padx=padx, pady=(0, 10), sticky="ew")
            self._quick_access_menus.append(menu)

        backup_card = ctk.CTkFrame(
            self.shell,
            corner_radius=16,
            fg_color=self.theme.PANEL_2,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        backup_card.grid(row=3, column=0, padx=14, pady=(0, 14), sticky="ew")
        backup_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            backup_card,
            text="Backups",
            font=self.font(16, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=0, column=0, padx=16, pady=(16, 6), sticky="w")

        ctk.CTkSwitch(
            backup_card,
            text="Backups aktivieren",
            variable=self._backup_enabled_var,
            onvalue=True,
            offvalue=False,
            text_color=self.theme.TEXT,
            command=self._persist_backup_preferences_safe,
        ).grid(row=1, column=0, padx=16, pady=(0, 12), sticky="w")

        ctk.CTkSwitch(
            backup_card,
            text="Automatisches Backup aktivieren",
            variable=self._auto_backup_enabled_var,
            onvalue=True,
            offvalue=False,
            text_color=self.theme.TEXT,
            command=self._persist_backup_preferences_safe,
        ).grid(row=2, column=0, padx=16, pady=(0, 12), sticky="w")

        ctk.CTkLabel(backup_card, text="Backup-Ordner", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=3, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        backup_dir_row = ctk.CTkFrame(backup_card, fg_color="transparent")
        backup_dir_row.grid(row=4, column=0, padx=16, pady=(0, 10), sticky="ew")
        backup_dir_row.grid_columnconfigure(0, weight=1)
        self.backup_dir_entry = ctk.CTkEntry(
            backup_dir_row,
            textvariable=self._backup_dir_var,
            height=38,
            corner_radius=12,
            placeholder_text="Backup-Ordner",
        )
        self.backup_dir_entry.grid(row=0, column=0, padx=(0, 8), sticky="ew")
        self.backup_dir_entry.bind("<FocusOut>", lambda _event: self._persist_backup_preferences_safe())
        ctk.CTkButton(
            backup_dir_row,
            text="Ordner wählen...",
            height=38,
            corner_radius=12,
            fg_color=self.theme.PANEL,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=self._select_backup_dir,
        ).grid(row=0, column=1, sticky="ew")

        settings_row = ctk.CTkFrame(backup_card, fg_color="transparent")
        settings_row.grid(row=5, column=0, padx=16, pady=(0, 10), sticky="ew")
        settings_row.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(settings_row, text="Standard-Backupmodus", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=0, column=0, padx=(0, 8), pady=(0, 6), sticky="w"
        )
        ctk.CTkLabel(settings_row, text="Aufbewahrung (Tage)", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=0, column=1, padx=(8, 0), pady=(0, 6), sticky="w"
        )

        self.backup_mode_menu = ctk.CTkOptionMenu(
            settings_row,
            values=["Vollbackup", "Teilbackup"],
            command=self._on_backup_mode_change,
            height=38,
            corner_radius=12,
            fg_color=self.theme.PANEL,
            button_color=self.theme.ACCENT,
            button_hover_color=self.theme.ACCENT_HOVER,
            text_color=self.theme.TEXT,
        )
        self.backup_mode_menu.grid(row=1, column=0, padx=(0, 8), sticky="ew")

        self.backup_retention_entry = ctk.CTkEntry(
            settings_row,
            textvariable=self._backup_retention_var,
            height=38,
            corner_radius=12,
            placeholder_text="30",
        )
        self.backup_retention_entry.grid(row=1, column=1, padx=(8, 0), sticky="ew")
        self.backup_retention_entry.bind("<FocusOut>", lambda _event: self._persist_backup_preferences_safe())

        auto_row = ctk.CTkFrame(backup_card, fg_color="transparent")
        auto_row.grid(row=6, column=0, padx=16, pady=(0, 10), sticky="ew")
        auto_row.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkLabel(auto_row, text="Automatisch alle X Tage", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=0, column=0, padx=(0, 8), pady=(0, 6), sticky="w"
        )
        ctk.CTkLabel(auto_row, text="Letztes Backup", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=0, column=1, padx=(8, 0), pady=(0, 6), sticky="w"
        )

        self.auto_backup_interval_entry = ctk.CTkEntry(
            auto_row,
            textvariable=self._auto_backup_interval_var,
            height=38,
            corner_radius=12,
            placeholder_text="7",
        )
        self.auto_backup_interval_entry.grid(row=1, column=0, padx=(0, 8), sticky="ew")
        self.auto_backup_interval_entry.bind("<FocusOut>", lambda _event: self._persist_backup_preferences_safe())

        self.last_backup_info = ctk.CTkLabel(
            auto_row,
            text="-",
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
            anchor="w",
            justify="left",
        )
        self.last_backup_info.grid(row=1, column=1, padx=(8, 0), sticky="w")

        action_row = ctk.CTkFrame(backup_card, fg_color="transparent")
        action_row.grid(row=7, column=0, padx=16, pady=(0, 8), sticky="ew")
        action_row.grid_columnconfigure((0, 1), weight=1)

        self.backup_now_button = ctk.CTkButton(
            action_row,
            text="Jetzt Backup erstellen",
            height=40,
            corner_radius=12,
            fg_color=self.theme.ACCENT,
            hover_color=self.theme.ACCENT_HOVER,
            text_color=("white", "white"),
            command=self.on_backup_now,
        )
        self.backup_now_button.grid(row=0, column=0, padx=(0, 8), sticky="ew")

        self.restore_backup_button = ctk.CTkButton(
            action_row,
            text="Backup wiederherstellen...",
            height=40,
            corner_radius=12,
            fg_color=self.theme.PANEL,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=self.on_restore_backup,
        )
        self.restore_backup_button.grid(row=0, column=1, padx=(8, 0), sticky="ew")

        self.backup_progress = ctk.CTkProgressBar(backup_card, mode="indeterminate", height=10)
        self.backup_progress.grid(row=8, column=0, padx=16, pady=(0, 8), sticky="ew")
        self.backup_progress.stop()
        self.backup_progress.set(0)

        ctk.CTkLabel(
            backup_card,
            textvariable=self._backup_status_var,
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
            justify="left",
            wraplength=900,
        ).grid(row=9, column=0, padx=16, pady=(0, 16), sticky="w")

    def _settings(self) -> dict:
        return self.settings_manager.load()

    def _save_settings(self, extra_updates=None):
        settings = self.settings_manager.load()
        settings.update(extra_updates or {})
        saved = self.settings_manager.save(settings)
        self.app.xml_folder = saved.get("xml_folder") or None
        self.app.appearance_preference = saved.get("appearance_mode") or "System"
        self.app.quick_access_items = self.app.normalize_quick_access_items(saved.get("quick_access_items", []))
        if hasattr(self.app, "refresh_quick_access_tools"):
            self.app.refresh_quick_access_tools()
        return saved

    def _quick_access_id_to_label(self, item_id: str) -> str:
        return self._quick_access_option_map.get(str(item_id or "").strip(), "Kein Eintrag")

    def _quick_access_label_to_id(self, label: str) -> str:
        selected = str(label or "").strip()
        for item_id, item_label in self.app.get_quick_access_options():
            if item_label == selected:
                return item_id
        return ""

    def _on_quick_access_change(self, _slot_index: int):
        items = [self._quick_access_label_to_id(variable.get()) for variable in self._quick_access_vars]
        normalized = self.app.normalize_quick_access_items(items)
        self._save_settings({"quick_access_items": normalized})
        for variable, menu, item_id in zip(self._quick_access_vars, self._quick_access_menus, normalized):
            menu.set(self._quick_access_id_to_label(item_id))
        self.refresh()

    def _set_appearance(self, mode: str):
        saved = self._save_settings({"appearance_mode": mode})
        self.app.set_appearance_preference(saved.get("appearance_mode"), persist=False)
        self.refresh()

    def _select_xml_folder_and_refresh(self):
        folder = filedialog.askdirectory(title="Ordner mit XML-Dateien auswählen")
        if not folder:
            return
        self._save_settings({"xml_folder": folder})
        try:
            self.app._sync_seen_files()
        except Exception:
            pass
        self.refresh()
        messagebox.showinfo("Ordner gespeichert", f"XML-Ordner:\n{folder}")

    def _select_backup_dir(self):
        folder = filedialog.askdirectory(title="Backup-Ordner auswählen")
        if folder:
            self._backup_dir_var.set(folder)
            self._persist_backup_preferences_safe()

    def _on_backup_mode_change(self, value):
        self._backup_mode_var.set("incremental" if value == "Teilbackup" else "full")
        self._persist_backup_preferences_safe()

    def _mode_label(self, mode: str) -> str:
        return "Teilbackup" if str(mode).strip().lower() == "incremental" else "Vollbackup"

    def _backup_mode_dialog(self, default_mode: str) -> str | None:
        result = {"mode": None}
        dialog = ctk.CTkToplevel(self)
        dialog.title("Backupmodus wählen")
        dialog.geometry("360x220")
        dialog.resizable(False, False)
        dialog.configure(fg_color=self.theme.BG)
        dialog.attributes("-topmost", True)

        shell = ctk.CTkFrame(dialog, corner_radius=18, fg_color=self.theme.PANEL, border_width=1, border_color=self.theme.BORDER)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        shell.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(shell, text="Backupmodus wählen", font=self.font(16, "bold"), text_color=self.theme.TEXT).grid(
            row=0, column=0, padx=16, pady=(16, 8), sticky="w"
        )
        ctk.CTkLabel(
            shell,
            text=f"Standard: {self._mode_label(default_mode)}",
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
        ).grid(row=1, column=0, padx=16, pady=(0, 12), sticky="w")

        button_row = ctk.CTkFrame(shell, fg_color="transparent")
        button_row.grid(row=2, column=0, padx=16, pady=(0, 16), sticky="ew")
        button_row.grid_columnconfigure((0, 1), weight=1)

        def _choose(mode: str):
            result["mode"] = mode
            dialog.destroy()

        ctk.CTkButton(
            button_row,
            text="Vollbackup",
            height=42,
            corner_radius=12,
            fg_color=self.theme.ACCENT if default_mode == "full" else self.theme.PANEL,
            hover_color=self.theme.ACCENT_HOVER if default_mode == "full" else self.theme.BORDER,
            text_color=("white", "white") if default_mode == "full" else self.theme.TEXT,
            command=lambda: _choose("full"),
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            button_row,
            text="Teilbackup",
            height=42,
            corner_radius=12,
            fg_color=self.theme.ACCENT if default_mode == "incremental" else self.theme.PANEL,
            hover_color=self.theme.ACCENT_HOVER if default_mode == "incremental" else self.theme.BORDER,
            text_color=("white", "white") if default_mode == "incremental" else self.theme.TEXT,
            command=lambda: _choose("incremental"),
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        ctk.CTkButton(
            shell,
            text="Abbrechen",
            height=38,
            corner_radius=12,
            fg_color=self.theme.PANEL,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=dialog.destroy,
        ).grid(row=3, column=0, padx=16, pady=(0, 16), sticky="ew")

        dialog.grab_set()
        dialog.focus_force()
        dialog.wait_window()
        return result["mode"]

    def _collect_backup_settings(self) -> dict:
        try:
            retention = int(self._backup_retention_var.get().strip())
        except Exception as exc:
            raise ValueError("Aufbewahrung (Tage) muss eine Zahl sein.") from exc

        if retention < 1 or retention > 365:
            raise ValueError("Aufbewahrung (Tage) muss zwischen 1 und 365 liegen.")

        try:
            auto_interval = int(self._auto_backup_interval_var.get().strip())
        except Exception as exc:
            raise ValueError("Automatisches Backup-Intervall muss eine Zahl sein.") from exc

        if auto_interval < 1 or auto_interval > 365:
            raise ValueError("Automatisches Backup-Intervall muss zwischen 1 und 365 Tagen liegen.")

        raw_backup_dir = self._backup_dir_var.get().strip()
        if not raw_backup_dir:
            raise ValueError("Bitte einen Backup-Ordner wählen.")
        backup_dir = Path(raw_backup_dir).expanduser()

        try:
            backup_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            raise ValueError(f"Backup-Ordner ist nicht verwendbar:\n{backup_dir}") from exc

        return self._save_settings(
            {
                "backups_enabled": bool(self._backup_enabled_var.get()),
                "auto_backup_enabled": bool(self._auto_backup_enabled_var.get()),
                "backup_dir": str(backup_dir),
                "backup_mode_default": self._backup_mode_var.get().strip().lower() or "full",
                "backup_retention_days": retention,
                "auto_backup_interval_days": auto_interval,
            }
        )

    def _persist_backup_preferences_safe(self):
        try:
            self._collect_backup_settings()
            if hasattr(self.app, "_check_auto_backup_due"):
                self.app.after(100, self.app._check_auto_backup_due)
        except Exception:
            return

    def _validate_backup_inputs(self) -> dict:
        return self._collect_backup_settings()

    def on_backup_now(self):
        if self._backup_running:
            return

        try:
            settings = self._validate_backup_inputs()
        except ValueError as exc:
            messagebox.showwarning("Backups", str(exc))
            return

        if not settings.get("backups_enabled"):
            if not messagebox.askyesno("Backups", "Backups sind deaktiviert. Trotzdem jetzt ein Backup erstellen?"):
                return

        mode = self._backup_mode_dialog(settings.get("backup_mode_default", "full"))
        if mode is None:
            return

        manager = BackupManager(
            app_name=str(getattr(self.app, "title", lambda: "GAWELA Tourenplaner")()).strip() or "GAWELA Tourenplaner",
            config_dir=Path(getattr(self.app, "config_dir", getattr(self.app, "base_dir", "."))),
            data_dir=Path(getattr(self.app, "data_dir", ".")),
            backup_dir=Path(settings["backup_dir"]),
        )

        self._backup_running = True
        self.backup_now_button.configure(state="disabled")
        self.backup_progress.start()
        self._backup_status_var.set(f"Backup läuft ({self._mode_label(mode)}) ...")

        def _worker():
            try:
                backup_path = manager.create_backup(mode)
                manager.cleanup_old_backups(settings["backup_retention_days"])
                self.after(0, lambda: self._on_backup_success(backup_path))
            except Exception as exc:
                self.after(0, lambda: self._on_backup_error(exc))

        threading.Thread(target=_worker, daemon=True).start()

    def on_restore_backup(self):
        if self._backup_running:
            return

        raw_backup_dir = self._backup_dir_var.get().strip()
        if not raw_backup_dir:
            messagebox.showwarning("Backups", "Bitte zuerst einen Backup-Ordner wählen.")
            return

        backup_dir = Path(raw_backup_dir).expanduser()
        try:
            backup_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Backups", f"Backup-Ordner ist nicht verwendbar:\n{exc}")
            return

        bak_path = filedialog.askopenfilename(
            title="Backup auswählen",
            initialdir=str(backup_dir),
            filetypes=[("Backup-Dateien", "*.bak"), ("Alle Dateien", "*.*")],
        )
        if not bak_path:
            return

        selected_groups = self._restore_selection_dialog()
        if selected_groups is None:
            return

        if not messagebox.askyesno(
            "Backup wiederherstellen",
            "Die aktuellen Daten werden mit dem ausgewählten Backup überschrieben.\n"
            "Bitte führe währenddessen keine parallelen Änderungen in der App aus.\n\n"
            "Backup jetzt wiederherstellen?",
        ):
            return

        manager = BackupManager(
            app_name=str(getattr(self.app, "title", lambda: "GAWELA Tourenplaner")()).strip() or "GAWELA Tourenplaner",
            config_dir=Path(getattr(self.app, "config_dir", getattr(self.app, "base_dir", "."))),
            data_dir=Path(getattr(self.app, "data_dir", ".")),
            backup_dir=backup_dir,
        )

        self._backup_running = True
        self.backup_now_button.configure(state="disabled")
        self.restore_backup_button.configure(state="disabled")
        self.backup_progress.start()
        self._backup_status_var.set(f"Backup wird wiederhergestellt: {Path(bak_path).name}")

        def _worker():
            try:
                manager.restore_backup(
                    bak_path=Path(bak_path),
                    target_data_dir=Path(getattr(self.app, "data_dir", ".")),
                    target_config_dir=Path(getattr(self.app, "config_dir", getattr(self.app, "base_dir", "."))),
                    selected_groups=selected_groups,
                )
                self.after(0, lambda: self._on_restore_success(Path(bak_path), backup_dir, selected_groups))
            except Exception as exc:
                self.after(0, lambda: self._on_backup_error(exc, operation="Wiederherstellung"))

        threading.Thread(target=_worker, daemon=True).start()

    def _restore_selection_dialog(self):
        result = {"groups": None}
        dialog = ctk.CTkToplevel(self)
        dialog.title("Wiederherstellung auswählen")
        dialog.geometry("460x520")
        dialog.resizable(False, False)
        dialog.configure(fg_color=self.theme.BG)
        dialog.attributes("-topmost", True)

        shell = ctk.CTkFrame(
            dialog,
            corner_radius=18,
            fg_color=self.theme.PANEL,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        shell.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            shell,
            text="Welche Daten sollen wiederhergestellt werden?",
            font=self.font(16, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=0, column=0, padx=16, pady=(16, 8), sticky="w")

        ctk.CTkLabel(
            shell,
            text="Du kannst alle Daten oder nur bestimmte Bereiche aus dem Backup zurückspielen.",
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
            wraplength=400,
            justify="left",
        ).grid(row=1, column=0, padx=16, pady=(0, 12), sticky="w")

        select_all_var = tk.BooleanVar(value=True)
        group_vars = {key: tk.BooleanVar(value=True) for key, _label, _hint in self.RESTORE_GROUPS}

        list_shell = ctk.CTkScrollableFrame(
            shell,
            height=280,
            corner_radius=12,
            fg_color=self.theme.PANEL_2,
            scrollbar_fg_color=self.theme.SCROLLBAR_TRACK,
            scrollbar_button_color=self.theme.SCROLLBAR_BUTTON,
            scrollbar_button_hover_color=self.theme.SCROLLBAR_HOVER,
        )
        list_shell.grid(row=2, column=0, padx=16, pady=(0, 12), sticky="ew")
        list_shell.grid_columnconfigure(0, weight=1)

        def _sync_all(*_args):
            all_selected = all(var.get() for var in group_vars.values())
            select_all_var.set(all_selected)

        def _toggle_all():
            selected = bool(select_all_var.get())
            for var in group_vars.values():
                var.set(selected)

        ctk.CTkCheckBox(
            shell,
            text="Alle Daten wiederherstellen",
            variable=select_all_var,
            command=_toggle_all,
            text_color=self.theme.TEXT,
        ).grid(row=3, column=0, padx=16, pady=(0, 8), sticky="w")

        for row, (key, label, hint) in enumerate(self.RESTORE_GROUPS):
            item = ctk.CTkFrame(list_shell, corner_radius=12, fg_color=self.theme.PANEL)
            item.grid(row=row, column=0, padx=8, pady=6, sticky="ew")
            item.grid_columnconfigure(1, weight=1)

            ctk.CTkCheckBox(
                item,
                text="",
                width=24,
                variable=group_vars[key],
                command=_sync_all,
            ).grid(row=0, column=0, padx=(10, 8), pady=10, sticky="w")

            text_shell = ctk.CTkFrame(item, fg_color="transparent")
            text_shell.grid(row=0, column=1, padx=(0, 10), pady=8, sticky="ew")
            text_shell.grid_columnconfigure(0, weight=1)

            ctk.CTkLabel(
                text_shell,
                text=label,
                font=self.font(13, "bold"),
                text_color=self.theme.TEXT,
            ).grid(row=0, column=0, sticky="w")
            ctk.CTkLabel(
                text_shell,
                text=hint,
                font=self.font(11),
                text_color=self.theme.SUBTEXT,
            ).grid(row=1, column=0, sticky="w")

        buttons = ctk.CTkFrame(shell, fg_color="transparent")
        buttons.grid(row=4, column=0, padx=16, pady=(4, 16), sticky="ew")
        buttons.grid_columnconfigure((0, 1), weight=1)

        def _confirm():
            selected = [key for key, var in group_vars.items() if var.get()]
            if not selected:
                messagebox.showwarning("Wiederherstellung", "Bitte mindestens einen Datenbereich auswählen.", parent=dialog)
                return
            result["groups"] = selected
            dialog.destroy()

        ctk.CTkButton(
            buttons,
            text="Abbrechen",
            height=40,
            corner_radius=12,
            fg_color=self.theme.PANEL_2,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=dialog.destroy,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            buttons,
            text="Wiederherstellen",
            height=40,
            corner_radius=12,
            fg_color=self.theme.ACCENT,
            hover_color=self.theme.ACCENT_HOVER,
            text_color=("white", "white"),
            command=_confirm,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        dialog.grab_set()
        dialog.focus_force()
        dialog.wait_window()
        return result["groups"]

    def _on_backup_success(self, backup_path: Path):
        self._backup_running = False
        self.backup_progress.stop()
        self.backup_progress.set(0)
        self.backup_now_button.configure(state="normal")
        self.restore_backup_button.configure(state="normal")
        try:
            self._save_settings({"last_backup_iso": datetime.now(timezone.utc).isoformat()})
        except Exception:
            pass
        self._backup_status_var.set(f"Letztes Backup: {backup_path.name}")
        self.refresh()
        messagebox.showinfo("Backups", f"Backup erfolgreich erstellt:\n{backup_path}")

    def _on_restore_success(self, backup_path: Path, backup_dir: Path, selected_groups):
        self._backup_running = False
        self.backup_progress.stop()
        self.backup_progress.set(0)
        self.backup_now_button.configure(state="normal")
        self.restore_backup_button.configure(state="normal")
        self._backup_status_var.set(f"Wiederhergestellt: {backup_path.name}")

        selected = set(selected_groups or [])
        restore_all = not selected or "all" in selected

        try:
            if restore_all or "settings" in selected:
                self.app.load_config()
        except Exception:
            pass

        try:
            if restore_all or "orders" in selected:
                self.app.load_pins()
        except Exception:
            pass

        try:
            if restore_all or "tours" in selected:
                self.app._tours_cache = None
                if hasattr(self.app, "_invalidate_tour_data_caches"):
                    self.app._invalidate_tour_data_caches()
        except Exception:
            pass

        try:
            if restore_all or "employees" in selected:
                self.app._employees_cache = None
        except Exception:
            pass

        try:
            if restore_all or "vehicles" in selected:
                self.app._vehicle_data_cache = None
        except Exception:
            pass

        try:
            if restore_all or "misc" in selected:
                self.app._geocode_cache = self.app._load_geocode_cache()
                self.app._geocode_cache_dirty = False
        except Exception:
            pass

        try:
            if restore_all or "tours" in selected or "employees" in selected or "vehicles" in selected:
                self.app._refresh_tour_related_views()
        except Exception:
            pass

        try:
            if hasattr(self.app, "update_route_employee_summary"):
                self.app.update_route_employee_summary()
            if hasattr(self.app, "update_route_resource_summary"):
                self.app.update_route_resource_summary()
        except Exception:
            pass

        try:
            for page_name in ("list", "tours", "employees", "vehicles", "settings"):
                page = getattr(self.app, "pages", {}).get(page_name)
                if page and hasattr(page, "refresh"):
                    page.refresh()
        except Exception:
            pass
        self.refresh()
        selected_labels = [
            BackupManager.RESTORE_LABELS.get(key, key)
            for key in (selected_groups or [])
        ]
        messagebox.showinfo(
            "Backups",
            f"Backup erfolgreich wiederhergestellt:\n{backup_path}\n\n"
            f"Wiederhergestellt: {', '.join(selected_labels)}",
        )

    def _on_backup_error(self, exc: Exception, operation: str = "Backup"):
        self._backup_running = False
        self.backup_progress.stop()
        self.backup_progress.set(0)
        self.backup_now_button.configure(state="normal")
        self.restore_backup_button.configure(state="normal")
        self._backup_status_var.set(f"Fehler: {exc}")
        messagebox.showerror("Backups", f"{operation} konnte nicht abgeschlossen werden:\n{exc}")

    def refresh(self):
        settings = self._settings()

        current = settings.get("appearance_mode", self.app.get_appearance_preference())
        for mode, button in self._appearance_buttons.items():
            selected = mode == current
            button.configure(
                fg_color=self.theme.ACCENT if selected else self.theme.PANEL,
                hover_color=self.theme.ACCENT_HOVER if selected else self.theme.BORDER,
                text_color=("white", "white") if selected else self.theme.TEXT,
                border_color=self.theme.ACCENT if selected else self.theme.BORDER,
                border_width=1,
                font=self.font(13, "bold" if selected else "normal"),
            )

        if self._theme_info is not None:
            current_label = {"System": "System", "Light": "Hell", "Dark": "Dunkel"}.get(current, current)
            self._theme_info.configure(text=f"Aktuelle Darstellung: {current_label}")

        if self._xml_info is not None:
            folder = str(settings.get("xml_folder") or "").strip()
            self._xml_info.configure(text=f"XML-Ordner: {folder}" if folder else "XML-Ordner: Noch nicht gesetzt")

        quick_access_items = self.app.normalize_quick_access_items(settings.get("quick_access_items", []))
        for variable, menu, item_id in zip(self._quick_access_vars, self._quick_access_menus, quick_access_items):
            menu.set(self._quick_access_id_to_label(item_id))
        if self._quick_access_info is not None:
            labels = [self._quick_access_id_to_label(item_id) for item_id in quick_access_items if item_id]
            summary = ", ".join(labels) if labels else "Keine Schnellzugriffe ausgewählt"
            self._quick_access_info.configure(text=f"Aktive Einträge: {summary}")

        self._backup_enabled_var.set(bool(settings.get("backups_enabled", False)))
        self._auto_backup_enabled_var.set(bool(settings.get("auto_backup_enabled", False)))
        self._backup_dir_var.set(str(settings.get("backup_dir") or self.settings_manager.default_backup_dir()))
        self._backup_mode_var.set(str(settings.get("backup_mode_default") or "full"))
        self._backup_retention_var.set(str(settings.get("backup_retention_days", 30)))
        self._auto_backup_interval_var.set(str(settings.get("auto_backup_interval_days", 7)))
        self.backup_mode_menu.set(self._mode_label(self._backup_mode_var.get()))

        last_backup_iso = str(settings.get("last_backup_iso") or "").strip()
        last_backup_text = "Noch kein Backup gespeichert"
        if last_backup_iso:
            try:
                parsed = datetime.fromisoformat(last_backup_iso.replace("Z", "+00:00"))
                last_backup_text = parsed.astimezone().strftime("%d-%m-%Y %H:%M")
            except Exception:
                last_backup_text = last_backup_iso
        self.last_backup_info.configure(text=last_backup_text)

        latest_backup = BackupManager(
            app_name=str(getattr(self.app, "title", lambda: "GAWELA Tourenplaner")()).strip() or "GAWELA Tourenplaner",
            config_dir=Path(getattr(self.app, "config_dir", getattr(self.app, "base_dir", "."))),
            data_dir=Path(getattr(self.app, "data_dir", ".")),
            backup_dir=Path(self._backup_dir_var.get()),
        ).find_latest_backup()
        self._backup_status_var.set(
            f"Letztes Backup: {latest_backup.name}" if latest_backup else "Kein Backup vorhanden."
        )

    def on_show(self):
        self.refresh()
