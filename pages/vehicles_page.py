import tkinter as tk
from datetime import datetime
from tkinter import messagebox
from uuid import uuid4

import customtkinter as ctk


TAB_LABELS = {
    "vehicles": "Zugfahrzeuge",
    "trailers": "Anhänger",
}

FORM_KIND_LABELS = {
    "vehicles": "Zugfahrzeug",
    "trailers": "Anhänger",
}


def _scrollable_frame_kwargs(theme):
    return {
        "scrollbar_fg_color": theme.SCROLLBAR_TRACK,
        "scrollbar_button_color": theme.SCROLLBAR_BUTTON,
        "scrollbar_button_hover_color": theme.SCROLLBAR_HOVER,
    }


class VehiclesPage(ctk.CTkFrame):
    def __init__(self, master, app, theme, font_factory):
        super().__init__(master, fg_color=theme.BG)
        self.app = app
        self.theme = theme
        self.font = font_factory
        self.segment_value = tk.StringVar(value="vehicles")

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        topbar = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=self.theme.PANEL,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        topbar.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 12))
        topbar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(topbar, text="Fahrzeugverwaltung", font=self.font(18, "bold"), text_color=self.theme.TEXT).grid(
            row=0, column=0, padx=16, pady=14, sticky="w"
        )

        toggle_shell = ctk.CTkFrame(
            topbar,
            corner_radius=14,
            fg_color=self.theme.PANEL_2,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        toggle_shell.grid(row=0, column=1, padx=10, pady=10, sticky="e")
        toggle_shell.grid_columnconfigure((0, 1), weight=1)

        self.btn_show_vehicles = ctk.CTkButton(
            toggle_shell,
            text=TAB_LABELS["vehicles"],
            width=150,
            height=36,
            corner_radius=10,
            command=lambda: self._set_segment("vehicles"),
        )
        self.btn_show_vehicles.grid(row=0, column=0, padx=(4, 2), pady=4, sticky="ew")

        self.btn_show_trailers = ctk.CTkButton(
            toggle_shell,
            text=TAB_LABELS["trailers"],
            width=120,
            height=36,
            corner_radius=10,
            command=lambda: self._set_segment("trailers"),
        )
        self.btn_show_trailers.grid(row=0, column=1, padx=(2, 4), pady=4, sticky="ew")

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
        ).grid(row=0, column=2, padx=(8, 8), pady=10)

        self.add_button = ctk.CTkButton(
            topbar,
            text="+ Fahrzeug",
            height=36,
            corner_radius=12,
            fg_color=self.theme.ACCENT,
            hover_color=self.theme.ACCENT_HOVER,
            command=self.open_editor,
        )
        self.add_button.grid(row=0, column=3, padx=(0, 14), pady=10)

        self.info = ctk.CTkLabel(self, text="", anchor="w", text_color=self.theme.SUBTEXT, font=self.font(13))
        self.info.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 0))

        shell = ctk.CTkFrame(
            self,
            corner_radius=18,
            fg_color=self.theme.PANEL,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        shell.grid(row=2, column=0, sticky="nsew", padx=20, pady=(12, 20))
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(0, weight=1)

        self.scroll = ctk.CTkScrollableFrame(
            shell,
            corner_radius=16,
            fg_color=self.theme.PANEL_2,
            **_scrollable_frame_kwargs(self.theme),
        )
        self.scroll.grid(row=0, column=0, sticky="nsew", padx=14, pady=14)
        self.scroll.grid_columnconfigure(0, weight=1)

        self._render_segment_buttons()
        self.refresh()

    def on_show(self):
        self.refresh()

    def _current_kind(self) -> str:
        return self.segment_value.get() or "vehicles"

    def _set_segment(self, kind: str):
        self.segment_value.set(kind if kind in TAB_LABELS else "vehicles")
        self._render_segment_buttons()
        self.refresh()

    def _render_segment_buttons(self):
        current = self._current_kind()
        for kind, button in (
            ("vehicles", self.btn_show_vehicles),
            ("trailers", self.btn_show_trailers),
        ):
            selected = kind == current
            button.configure(
                fg_color="#800080" if selected else "transparent",
                hover_color="#660066" if selected else self.theme.BORDER,
                text_color=("white", "white") if selected else self.theme.TEXT,
                border_width=1 if selected else 0,
                border_color="#800080" if selected else self.theme.PANEL_2,
                font=self.font(13, "bold" if selected else "normal"),
            )

    def refresh(self):
        for child in self.scroll.winfo_children():
            child.destroy()

        kind = self._current_kind()
        dataset = self.app._load_vehicle_data()
        items = dataset.get(kind, [])
        is_vehicle = kind == "vehicles"

        self.add_button.configure(text="+ Fahrzeug" if is_vehicle else "+ Anhänger")
        self.info.configure(text=f"{TAB_LABELS[kind]}: {len(items)}")

        if not items:
            ctk.CTkLabel(
                self.scroll,
                text="Noch keine Einträge vorhanden.",
                font=self.font(13),
                text_color=self.theme.SUBTEXT,
            ).grid(row=0, column=0, padx=12, pady=12, sticky="w")
            return

        for row_idx, item in enumerate(items):
            self._build_card(row_idx, item, kind)

    def _build_card(self, row_idx: int, item: dict, kind: str):
        is_vehicle = kind == "vehicles"
        active = bool(item.get("active", True))
        card = ctk.CTkFrame(
            self.scroll,
            corner_radius=16,
            fg_color=self.theme.PANEL,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        card.grid(row=row_idx, column=0, padx=6, pady=6, sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        title_row = ctk.CTkFrame(card, fg_color="transparent")
        title_row.grid(row=0, column=0, padx=14, pady=(14, 6), sticky="ew")
        title_row.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            title_row,
            text=item.get("name", ""),
            font=self.font(15, "bold"),
            text_color=self.theme.TEXT,
        ).grid(row=0, column=0, sticky="w")

        status_text = "Aktiv" if active else "Inaktiv"
        status_color = self.theme.SUCCESS if active else self.theme.MUTED_BTN
        ctk.CTkLabel(
            title_row,
            text=status_text,
            font=self.font(11, "bold"),
            text_color=("white", "white"),
            fg_color=status_color,
            corner_radius=999,
            padx=10,
            pady=4,
        ).grid(row=0, column=1, padx=(8, 0), sticky="e")

        lines = []
        plate = str(item.get("license_plate") or "").strip() or "Kein Kennzeichen"
        lines.append(f"Kennzeichen: {plate}")
        lines.append(f"Nutzlast: {int(item.get('max_payload_kg', 0) or 0)} kg")
        if is_vehicle:
            lines.append(f"Anhängelast: {int(item.get('max_trailer_load_kg', 0) or 0)} kg")
        volume = int(item.get("volume_m3", 0) or 0)
        if volume > 0:
            lines.append(f"Volumen: {volume} m3")

        loading_area = item.get("loading_area") or {}
        if isinstance(loading_area, dict) and any(int(loading_area.get(key, 0) or 0) > 0 for key in ("length_cm", "width_cm", "height_cm")):
            lines.append(
                "Ladefläche: "
                f"{int(loading_area.get('length_cm', 0) or 0)} x "
                f"{int(loading_area.get('width_cm', 0) or 0)} x "
                f"{int(loading_area.get('height_cm', 0) or 0)} cm"
            )
        if str(item.get("notes") or "").strip():
            lines.append(f"Notizen: {item.get('notes')}")

        ctk.CTkLabel(
            card,
            text=" | ".join(lines),
            font=self.font(12),
            text_color=self.theme.SUBTEXT,
            justify="left",
            wraplength=900,
        ).grid(row=1, column=0, padx=14, pady=(0, 10), sticky="w")

        btns = ctk.CTkFrame(card, fg_color="transparent")
        btns.grid(row=2, column=0, padx=14, pady=(0, 14), sticky="ew")
        btns.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            btns,
            text="Bearbeiten",
            height=36,
            corner_radius=12,
            fg_color=self.theme.PANEL_2,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=lambda value=item, record_kind=kind: self.open_editor(value, record_kind),
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            btns,
            text="Löschen",
            height=36,
            corner_radius=12,
            fg_color=self.theme.DANGER,
            hover_color=self.theme.DANGER_HOVER,
            command=lambda value=item, vehicle=is_vehicle: self.delete_item(value, vehicle),
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

    def open_editor(self, item=None, source_kind=None):
        item = item if isinstance(item, dict) else {}
        source_kind = source_kind or self._current_kind()

        dlg = ctk.CTkToplevel(self)
        dlg.title("Fahrzeug")
        dlg.geometry("520x600")
        dlg.resizable(False, False)
        dlg.configure(fg_color=self.theme.BG)
        dlg.attributes("-topmost", True)

        shell = ctk.CTkFrame(dlg, corner_radius=18, fg_color=self.theme.PANEL, border_width=1, border_color=self.theme.BORDER)
        shell.pack(fill="both", expand=True, padx=16, pady=16)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(0, weight=1)

        form_scroll = ctk.CTkScrollableFrame(
            shell,
            corner_radius=14,
            fg_color=self.theme.PANEL,
            **_scrollable_frame_kwargs(self.theme),
        )
        form_scroll.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        form_scroll.grid_columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(form_scroll, text=FORM_KIND_LABELS[source_kind], font=self.font(16, "bold"), text_color=self.theme.TEXT)
        title_label.grid(row=0, column=0, padx=16, pady=(16, 10), sticky="w")

        kind_var = tk.StringVar(value=FORM_KIND_LABELS[source_kind])
        ctk.CTkLabel(form_scroll, text="Typ", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=1, column=0, padx=16, pady=(0, 6), sticky="w"
        )

        kind_switch_shell = ctk.CTkFrame(
            form_scroll,
            corner_radius=14,
            fg_color=self.theme.PANEL_2,
            border_width=1,
            border_color=self.theme.BORDER,
        )
        kind_switch_shell.grid(row=2, column=0, padx=16, pady=(0, 10), sticky="ew")
        kind_switch_shell.grid_columnconfigure((0, 1), weight=1)

        kind_buttons = {}

        def _selected_kind_key() -> str:
            return "trailers" if kind_var.get() == FORM_KIND_LABELS["trailers"] else "vehicles"

        def _render_kind_switch():
            current_kind = _selected_kind_key()
            for key, button in kind_buttons.items():
                selected = key == current_kind
                button.configure(
                    fg_color=self.theme.PANEL if selected else "transparent",
                    hover_color=self.theme.BORDER,
                    text_color=self.theme.TEXT,
                    border_width=2 if selected else 0,
                    border_color=self.theme.ACCENT if selected else self.theme.PANEL_2,
                    font=self.font(13, "bold" if selected else "normal"),
                )

        def _set_kind(kind_key: str):
            kind_var.set(FORM_KIND_LABELS[kind_key])
            _render_kind_switch()

        for column, kind_key in enumerate(("vehicles", "trailers")):
            button = ctk.CTkButton(
                kind_switch_shell,
                text=FORM_KIND_LABELS[kind_key],
                height=38,
                corner_radius=10,
                fg_color="transparent",
                hover_color=self.theme.BORDER,
                text_color=self.theme.TEXT,
                border_width=0,
                border_color=self.theme.PANEL_2,
                command=lambda value=kind_key: _set_kind(value),
            )
            button.grid(row=0, column=column, padx=(4, 2) if column == 0 else (2, 4), pady=4, sticky="ew")
            kind_buttons[kind_key] = button

        row_offset = 3
        ctk.CTkLabel(form_scroll, text="Name", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=row_offset, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        name_entry = ctk.CTkEntry(form_scroll, height=36, corner_radius=12, placeholder_text="Name *")
        name_entry.grid(row=row_offset + 1, column=0, padx=16, pady=(0, 10), sticky="ew")
        if str(item.get("name") or "").strip():
            name_entry.insert(0, item.get("name", ""))

        ctk.CTkLabel(form_scroll, text="Kennzeichen", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=row_offset + 2, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        plate_entry = ctk.CTkEntry(form_scroll, height=36, corner_radius=12, placeholder_text="Kennzeichen")
        plate_entry.grid(row=row_offset + 3, column=0, padx=16, pady=(0, 10), sticky="ew")
        if str(item.get("license_plate") or "").strip():
            plate_entry.insert(0, item.get("license_plate", ""))

        ctk.CTkLabel(form_scroll, text="Nutzlast (kg)", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=row_offset + 4, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        payload_entry = ctk.CTkEntry(form_scroll, height=36, corner_radius=12, placeholder_text="Nutzlast in kg")
        payload_entry.grid(row=row_offset + 5, column=0, padx=16, pady=(0, 10), sticky="ew")
        if int(item.get("max_payload_kg", 0) or 0) > 0:
            payload_entry.insert(0, str(item.get("max_payload_kg", 0)))

        trailer_load_label = ctk.CTkLabel(form_scroll, text="Anhängelast (kg)", font=self.font(12, "bold"), text_color=self.theme.TEXT)
        trailer_load_label.grid(row=row_offset + 6, column=0, padx=16, pady=(0, 6), sticky="w")
        trailer_load_entry = ctk.CTkEntry(form_scroll, height=36, corner_radius=12, placeholder_text="Anhängelast in kg")
        if int(item.get("max_trailer_load_kg", 0) or 0) > 0:
            trailer_load_entry.insert(0, str(item.get("max_trailer_load_kg", 0)))

        row_cursor = row_offset + 8
        ctk.CTkLabel(form_scroll, text="Volumen (m3)", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=row_cursor, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        volume_entry = ctk.CTkEntry(form_scroll, height=36, corner_radius=12, placeholder_text="Volumen in m3 (optional)")
        volume_entry.grid(row=row_cursor + 1, column=0, padx=16, pady=(0, 10), sticky="ew")
        if int(item.get("volume_m3", 0) or 0) > 0:
            volume_entry.insert(0, str(item.get("volume_m3", 0)))
        row_cursor += 2

        loading_area = item.get("loading_area") or {}
        ctk.CTkLabel(form_scroll, text="Ladefläche (L x B x H in cm)", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=row_cursor, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        dims = ctk.CTkFrame(form_scroll, fg_color="transparent")
        dims.grid(row=row_cursor + 1, column=0, padx=16, pady=(0, 10), sticky="ew")
        dims.grid_columnconfigure((0, 1, 2), weight=1)

        length_entry = ctk.CTkEntry(dims, height=36, corner_radius=12, placeholder_text="L cm")
        width_entry = ctk.CTkEntry(dims, height=36, corner_radius=12, placeholder_text="B cm")
        height_entry = ctk.CTkEntry(dims, height=36, corner_radius=12, placeholder_text="H cm")
        length_entry.grid(row=0, column=0, padx=(0, 6), sticky="ew")
        width_entry.grid(row=0, column=1, padx=6, sticky="ew")
        height_entry.grid(row=0, column=2, padx=(6, 0), sticky="ew")
        if int(loading_area.get("length_cm", 0) or 0) > 0:
            length_entry.insert(0, str(loading_area.get("length_cm", 0)))
        if int(loading_area.get("width_cm", 0) or 0) > 0:
            width_entry.insert(0, str(loading_area.get("width_cm", 0)))
        if int(loading_area.get("height_cm", 0) or 0) > 0:
            height_entry.insert(0, str(loading_area.get("height_cm", 0)))
        row_cursor += 2

        ctk.CTkLabel(form_scroll, text="Notizen", font=self.font(12, "bold"), text_color=self.theme.TEXT).grid(
            row=row_cursor, column=0, padx=16, pady=(0, 6), sticky="w"
        )
        notes_box = ctk.CTkTextbox(form_scroll, height=120, corner_radius=12, border_width=1, border_color=self.theme.BORDER)
        notes_box.grid(row=row_cursor + 1, column=0, padx=16, pady=(0, 10), sticky="ew")
        if str(item.get("notes") or "").strip():
            notes_box.insert("1.0", item.get("notes", ""))
        row_cursor += 2

        active_label = "Fahrzeug ist aktiv" if source_kind == "vehicles" else "Anhänger ist aktiv"
        active_var = tk.BooleanVar(value=item.get("active", True))
        active_checkbox = ctk.CTkCheckBox(
            form_scroll,
            text=active_label,
            variable=active_var,
            onvalue=True,
            offvalue=False,
            text_color=self.theme.TEXT,
        )
        active_checkbox.grid(row=row_cursor, column=0, padx=16, pady=(0, 12), sticky="w")
        row_cursor += 1

        btns = ctk.CTkFrame(form_scroll, fg_color="transparent")
        btns.grid(row=row_cursor, column=0, padx=16, pady=(6, 16), sticky="ew")
        btns.grid_columnconfigure((0, 1), weight=1)

        def _toggle_kind_fields(*_args):
            current_kind = _selected_kind_key()
            title_label.configure(text=FORM_KIND_LABELS[current_kind])
            active_checkbox.configure(text="Fahrzeug ist aktiv" if current_kind == "vehicles" else "Anhänger ist aktiv")
            if current_kind == "vehicles":
                trailer_load_label.grid(row=row_offset + 6, column=0, padx=16, pady=(0, 6), sticky="w")
                trailer_load_entry.grid(row=row_offset + 7, column=0, padx=16, pady=(0, 10), sticky="ew")
            else:
                trailer_load_label.grid_forget()
                trailer_load_entry.grid_forget()
            _render_kind_switch()

        kind_var.trace_add("write", _toggle_kind_fields)
        _toggle_kind_fields()

        def _save():
            created_at = item.get("created_at") or datetime.now().replace(microsecond=0).isoformat()
            target_kind = _selected_kind_key()
            payload = {
                "id": item.get("id") or str(uuid4()),
                "name": name_entry.get().strip(),
                "license_plate": plate_entry.get().strip(),
                "max_payload_kg": payload_entry.get().strip(),
                "active": bool(active_var.get()),
                "notes": notes_box.get("1.0", "end").strip(),
                "volume_m3": volume_entry.get().strip(),
                "loading_area": {
                    "length_cm": length_entry.get().strip() or 0,
                    "width_cm": width_entry.get().strip() or 0,
                    "height_cm": height_entry.get().strip() or 0,
                },
                "created_at": created_at,
            }

            try:
                if target_kind == "vehicles":
                    payload["type"] = str(item.get("type") or "other").strip() or "other"
                    payload["max_trailer_load_kg"] = trailer_load_entry.get().strip()
                    if source_kind == "trailers" and item.get("id"):
                        self.app.delete_trailer_record(item.get("id"), suppress_in_use_check=True)
                    self.app.upsert_vehicle_record(payload)
                else:
                    if source_kind == "vehicles" and item.get("id"):
                        self.app.delete_vehicle_record(item.get("id"), suppress_in_use_check=True)
                    self.app.upsert_trailer_record(payload)
            except ValueError as exc:
                messagebox.showwarning("Fahrzeug", str(exc))
                return
            except Exception as exc:
                messagebox.showerror("Fahrzeug", f"Eintrag konnte nicht gespeichert werden:\n{exc}")
                return

            self._set_segment(target_kind)
            try:
                dlg.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            btns,
            text="Abbrechen",
            height=40,
            corner_radius=14,
            fg_color=self.theme.PANEL_2,
            hover_color=self.theme.BORDER,
            text_color=self.theme.TEXT,
            command=dlg.destroy,
        ).grid(row=0, column=0, padx=(0, 8), sticky="ew")

        ctk.CTkButton(
            btns,
            text="Speichern",
            height=40,
            corner_radius=14,
            font=self.font(13, "bold"),
            fg_color=self.theme.SUCCESS,
            hover_color=self.theme.SUCCESS_HOVER,
            command=_save,
        ).grid(row=0, column=1, padx=(8, 0), sticky="ew")

        dlg.grab_set()
        dlg.focus_force()

    def delete_item(self, item: dict, is_vehicle: bool):
        label = item.get("name") or ("Fahrzeug" if is_vehicle else "Anhänger")
        if not messagebox.askyesno("Löschen", f"{label} wirklich löschen?"):
            return
        try:
            if is_vehicle:
                self.app.delete_vehicle_record(item.get("id"))
            else:
                self.app.delete_trailer_record(item.get("id"))
        except Exception as exc:
            messagebox.showerror("Fahrzeugverwaltung", f"Eintrag konnte nicht gelöscht werden:\n{exc}")
            return
        self.refresh()
