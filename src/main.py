"""
LEK-Bastler - Hauptprogramm v1.0
Grafische Benutzeroberfläche für die Aufgabenauswahl aus Word-Dokumenten

Features:
- Automatische Platzhalter-Ersetzung in LEK-Vorlagen (TitelFürThema → LEK-Thema)
- Korrekte Überschrift-Level-Erhaltung beim Kopieren von Aufgaben
- Vollständige Tabellen-Unterstützung mit Formatierung
- Professionelle LEK-Erstellung aus Word-Dokumenten
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import simpledialog
from tkinter import scrolledtext
from typing import Any
import json
import os
import re
import sys
from pathlib import Path
from word_processor import WordProcessor
from task_selector import TaskSelector
from import_wizard import ImportSession, is_blocking_warning


def _runtime_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


class HoverTooltip:
    """Zeigt einen kurzen Infotext an, wenn die Maus über einem Widget verweilt."""

    def __init__(self, widget, text, delay_ms=350):
        self.widget = widget
        self.text = str(text or '').strip()
        self.delay_ms = int(delay_ms)
        self._after_id = None
        self._tip_window = None

        self.widget.bind('<Enter>', self._on_enter, add='+')
        self.widget.bind('<Leave>', self._on_leave, add='+')
        self.widget.bind('<Motion>', self._on_motion, add='+')
        self.widget.bind('<ButtonPress>', self._on_leave, add='+')

    def _on_enter(self, _event=None):
        self._schedule_show()

    def _on_motion(self, _event=None):
        if self._tip_window is not None:
            self._position_tip()

    def _on_leave(self, _event=None):
        self._cancel_show()
        self._hide_tip()

    def _schedule_show(self):
        self._cancel_show()
        if not self.text:
            return
        self._after_id = self.widget.after(self.delay_ms, self._show_tip)

    def _cancel_show(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None

    def _show_tip(self):
        self._after_id = None
        if self._tip_window is not None or not self.text:
            return

        tip = tk.Toplevel(self.widget)
        tip.wm_overrideredirect(True)
        tip.attributes('-topmost', True)

        label = tk.Label(
            tip,
            text=self.text,
            justify=tk.LEFT,
            background='#ffffe1',
            relief=tk.SOLID,
            borderwidth=1,
            padx=8,
            pady=4,
            font=('Segoe UI', 9),
        )
        label.pack()

        self._tip_window = tip
        self._position_tip()

    def _position_tip(self):
        if self._tip_window is None:
            return

        try:
            x = self.widget.winfo_pointerx() + 14
            y = self.widget.winfo_pointery() + 14
            self._tip_window.wm_geometry(f'+{x}+{y}')
        except Exception:
            pass

    def _hide_tip(self):
        if self._tip_window is not None:
            try:
                self._tip_window.destroy()
            except Exception:
                pass
            self._tip_window = None

class LEKBastlerGUI:
    def __init__(self, root):
        self.root = root
        self._tooltips = []
        self._apply_window_icon()
        app_version = self._read_app_version()
        title_suffix = f" v{app_version}" if app_version else ""
        self.root.title(f"LEK-Bastler{title_suffix} - Aufgabenauswahl")
        self.root.geometry("1120x760")
        self.root.minsize(1024, 700)
        self._apply_window_state()
        
        self.word_processor = WordProcessor()
        self.task_selector = TaskSelector()
        self.loaded_tasks = []
        self.current_displayed_tasks = []  # Speichert die aktuell angezeigten Aufgaben
        self.source_filename = ""  # Speichert den Namen der Quelldatei
        self.lek_theme = ""  # Speichert das extrahierte LEK-Thema
        self.import_session = None  # Wizard-Session (Sprint 1)
        self._authoring_help_prompted = False
        self._sort_column = None
        self._sort_reverse = False
        self.ui_mode_var = tk.StringVar(value="einfach (empfohlen)")
        self.current_step = 1
        self._action_groups = []
        self._action_group_base_width = 560
        self._action_group_min_width = 360
        self._resize_after_id = None
        self.step_labels = {
            1: "Quelle wählen",
            2: "Erkennung prüfen",
            3: "Aufgaben freigeben",
            4: "Exportieren",
        }
        self._default_button_labels = {}
        
        self.setup_ui()

    def _apply_window_state(self):
        """Startet das Hauptfenster im normalen Fensterzustand."""
        try:
            self.root.state('normal')
        except Exception:
            try:
                self.root.wm_state('normal')
            except Exception:
                pass

    def _apply_window_icon(self):
        """Setzt das Fenster-Icon (Titelleiste) wenn verfügbar."""
        try:
            if getattr(sys, 'frozen', False):
                base_dir = Path(getattr(sys, '_MEIPASS', os.path.dirname(sys.executable)))
            else:
                base_dir = Path(os.path.dirname(os.path.abspath(__file__)))

            icon_path = base_dir / "app_icon.ico"
            if icon_path.exists():
                self.root.iconbitmap(str(icon_path))
        except Exception:
            # Kein harter Fehler, falls Icon in einer Umgebung nicht gesetzt werden kann
            pass

    def _read_app_version(self):
        """Liest die App-Version aus build_version_info.txt (falls vorhanden)."""
        candidates = [
            _runtime_base_dir() / "src" / "build_version_info.txt",
            _runtime_base_dir() / "build_version_info.txt",
        ]

        version_file = next((path for path in candidates if path.exists()), None)
        if version_file is not None:
            try:
                content = version_file.read_text(encoding='utf-8', errors='ignore')
            except Exception:
                content = ""

            if content:
                m = re.search(r"StringStruct\('FileVersion',\s*'([^']+)'\)", content)
                if m:
                    return str(m.group(1)).strip()

                tuple_match = re.search(r"filevers\s*=\s*\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*(\d+))?\)", content)
                if tuple_match:
                    major, minor, patch, build = tuple_match.groups()
                    if build and str(build).strip() not in ('', '0'):
                        return f"{major}.{minor}.{patch}.{build}"
                    return f"{major}.{minor}.{patch}"

        # Fallback für Deploy/EXE-Betrieb: Version aus Dateinamen/Ordnernamen ermitteln
        # Beispiele: LEK-Bastler_3.7.0.exe, LEK-Bastler_3.7.0
        runtime_markers = []
        try:
            exe_path = Path(sys.executable).resolve()
            runtime_markers.extend([
                exe_path.name,
                exe_path.stem,
                exe_path.parent.name,
            ])
        except Exception:
            pass

        try:
            runtime_markers.append(_runtime_base_dir().name)
        except Exception:
            pass

        try:
            runtime_markers.append(Path(__file__).resolve().parent.name)
        except Exception:
            pass

        try:
            arg0 = Path(str(sys.argv[0] or '')).name
            if arg0:
                runtime_markers.append(arg0)
        except Exception:
            pass

        for marker in runtime_markers:
            text = str(marker or '').strip()
            if not text:
                continue

            explicit = re.search(r'LEK-Bastler[_-](\d+\.\d+\.\d+(?:\.\d+)?)', text, flags=re.IGNORECASE)
            if explicit:
                return str(explicit.group(1)).strip()

            generic = re.search(r'(\d+\.\d+\.\d+(?:\.\d+)?)', text)
            if generic:
                return str(generic.group(1)).strip()

        return ""
    
    def setup_ui(self):
        """Erstellt die Benutzeroberfläche"""
        
        # Hauptframe
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="wens")
        
        # Datei-Auswahl Sektion
        file_frame = ttk.LabelFrame(main_frame, text="Aufgabensammlung laden", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky="we", pady=(0, 10))

        ttk.Label(file_frame, text="Modus:").grid(row=0, column=3, sticky=tk.E, padx=(10, 0))
        self.ui_mode_combo = ttk.Combobox(
            file_frame,
            textvariable=self.ui_mode_var,
            values=["einfach (empfohlen)", "erweitert"],
            state="readonly",
            width=12,
        )
        self.ui_mode_combo.grid(row=0, column=4, sticky=tk.W, padx=(8, 0))
        self.ui_mode_combo.bind("<<ComboboxSelected>>", self._on_ui_mode_changed)
        self.ui_mode_hint_var = tk.StringVar(value="Einfach blendet Expertenfunktionen aus.")
        ttk.Label(file_frame, textvariable=self.ui_mode_hint_var).grid(row=0, column=5, sticky=tk.W, padx=(10, 0))
        
        self.file_path_var = tk.StringVar()
        ttk.Label(file_frame, text="Word-Datei:").grid(row=0, column=0, sticky=tk.W)
        self.entry_file_path = ttk.Entry(file_frame, textvariable=self.file_path_var, width=50)
        self.entry_file_path.grid(row=0, column=1, padx=(10, 0))
        self.btn_browse = ttk.Button(file_frame, text="Datei wählen & laden...", command=self.browse_and_load_file)
        self.btn_browse.grid(row=0, column=2, padx=(10, 0))
        self.btn_import_task = ttk.Button(
            file_frame,
            text="Aufgabe aus Word übernehmen...",
            command=self.import_task_from_word,
        )
        self.btn_import_task.grid(row=1, column=1, sticky=tk.W, pady=(8, 0))
        self.btn_import_tasks_bulk = ttk.Button(
            file_frame,
            text="Mehrere Aufgaben aus Word übernehmen...",
            command=self.import_tasks_bulk_from_word,
        )
        self.btn_import_tasks_bulk.grid(row=1, column=2, sticky=tk.W, padx=(10, 0), pady=(8, 0))
        self.btn_authoring_help = ttk.Button(
            file_frame,
            text="Eingabehilfe (JSON & Formulierungen)",
            command=self.show_task_authoring_help,
        )
        self.btn_authoring_help.grid(row=1, column=4, sticky=tk.W, padx=(10, 0), pady=(8, 0))

        self.lbl_duplicate_mode = ttk.Label(file_frame, text="Duplikat-Preset:")
        self.lbl_duplicate_mode.grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
        self.duplicate_mode_var = tk.StringVar()
        self.duplicate_mode_combo = ttk.Combobox(
            file_frame,
            textvariable=self.duplicate_mode_var,
            values=self._available_duplicate_modes(),
            state="readonly",
            width=18,
        )
        self.duplicate_mode_combo.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=(8, 0))
        self.duplicate_mode_combo.bind("<<ComboboxSelected>>", lambda _e: self._apply_duplicate_mode_from_ui(show_message=False))

        self.btn_apply_duplicate_mode = ttk.Button(
            file_frame,
            text="Preset anwenden",
            command=self._apply_duplicate_mode_from_ui,
        )
        self.btn_apply_duplicate_mode.grid(row=2, column=2, sticky=tk.W, padx=(10, 0), pady=(8, 0))

        self.btn_save_duplicate_mode = ttk.Button(
            file_frame,
            text="Als Standard speichern",
            command=self._save_duplicate_mode_as_default,
        )
        self.btn_save_duplicate_mode.grid(row=2, column=3, sticky=tk.W, padx=(10, 0), pady=(8, 0))

        self.duplicate_threshold_var = tk.StringVar(value="Aktive Duplikat-Schwelle: -")
        self.lbl_duplicate_threshold = ttk.Label(file_frame, textvariable=self.duplicate_threshold_var)
        self.lbl_duplicate_threshold.grid(
            row=3,
            column=1,
            columnspan=2,
            sticky=tk.W,
            padx=(10, 0),
            pady=(4, 0),
        )
        
        # Kriterien-Auswahl Sektion
        criteria_frame = ttk.LabelFrame(main_frame, text="Auswahlkriterien", padding="10")
        criteria_frame.grid(row=1, column=0, columnspan=2, sticky="we", pady=(0, 10))
        
        # Suchbegriffe
        ttk.Label(criteria_frame, text="Suchbegriffe (kommagetrennt):").grid(row=0, column=0, sticky=tk.W)
        self.keywords_var = tk.StringVar()
        self.entry_keywords_filter = ttk.Entry(criteria_frame, textvariable=self.keywords_var, width=60)
        self.entry_keywords_filter.grid(row=0, column=1, padx=(10, 0))
        
        # Schwierigkeitsgrad
        ttk.Label(criteria_frame, text="Schwierigkeitsgrad:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.difficulty_var = tk.StringVar()
        self.combo_filter_difficulty = ttk.Combobox(
            criteria_frame,
            textvariable=self.difficulty_var,
            values=["Alle", "Leicht", "Mittel", "Schwer"],
        )
        self.combo_filter_difficulty.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=(10, 0))
        self.combo_filter_difficulty.set("Alle")
        
        # Anzahl der Aufgaben
        ttk.Label(criteria_frame, text="Max. Anzahl Aufgaben:").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        self.max_tasks_var = tk.StringVar(value="20")
        self.spin_max_tasks = ttk.Spinbox(criteria_frame, from_=1, to=20, textvariable=self.max_tasks_var, width=10)
        self.spin_max_tasks.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=(10, 0))

        self.only_blockers_var = tk.BooleanVar(value=False)
        self.chk_only_blockers = ttk.Checkbutton(
            criteria_frame,
            text="Nur Blocker anzeigen",
            variable=self.only_blockers_var,
            command=self._on_toggle_only_blockers,
        )
        self.chk_only_blockers.grid(row=3, column=1, sticky=tk.W, padx=(10, 0), pady=(10, 0))
        
        # Wizard-Navigation
        wizard_frame = ttk.LabelFrame(main_frame, text="Import-Assistent", padding="10")
        wizard_frame.grid(row=2, column=0, columnspan=2, sticky="we", pady=(0, 10))

        self.wizard_step_var = tk.StringVar(value="Schritt 1/4: Quelle wählen")
        ttk.Label(wizard_frame, textvariable=self.wizard_step_var).grid(row=0, column=0, sticky=tk.W)

        self.wizard_hint_var = tk.StringVar(value="Nächster Schritt: Aufgabensammlung auswählen und laden.")
        ttk.Label(wizard_frame, textvariable=self.wizard_hint_var).grid(row=1, column=0, sticky=tk.W, pady=(4, 0))

        nav_frame = ttk.Frame(wizard_frame)
        nav_frame.grid(row=0, column=1, rowspan=2, sticky=tk.E)
        self.btn_step_prev = ttk.Button(nav_frame, text="← Zurück", command=self._on_prev_step)
        self.btn_step_prev.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_step_next = ttk.Button(nav_frame, text="Weiter →", command=self._on_next_step)
        self.btn_step_next.pack(side=tk.LEFT)

        wizard_frame.columnconfigure(0, weight=1)

        # Aufgaben-Vorschau
        preview_frame = ttk.LabelFrame(main_frame, text="Gefundene Aufgaben", padding="10")
        preview_frame.grid(row=3, column=0, columnspan=2, sticky="wens", pady=(0, 10))

        self.wizard_status_var = tk.StringVar(value="Wizard-Status: Keine Datei geladen")
        ttk.Label(preview_frame, textvariable=self.wizard_status_var).grid(
            row=0, column=0, sticky="we", pady=(0, 6)
        )
        
        # Treeview für Aufgabenvorschau
        columns = ("Nr", "Kategorie", "Titel", "Schwierigkeit", "Hinweise", "Suchbegriffe")
        self.task_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=10, selectmode="extended")
        
        # Spaltenbreiten optimiert setzen
        self.task_tree.heading("Nr", text="Nr", command=lambda: self._sort_task_tree_by_column("Nr"))
        self.task_tree.column("Nr", width=50, minwidth=40)
        
        self.task_tree.heading("Kategorie", text="Kategorie", command=lambda: self._sort_task_tree_by_column("Kategorie"))
        self.task_tree.column("Kategorie", width=170, minwidth=120)

        self.task_tree.heading("Titel", text="Titel", command=lambda: self._sort_task_tree_by_column("Titel"))
        self.task_tree.column("Titel", width=260, minwidth=180)
        
        self.task_tree.heading("Schwierigkeit", text="Schwierigkeit", command=lambda: self._sort_task_tree_by_column("Schwierigkeit"))
        self.task_tree.column("Schwierigkeit", width=100, minwidth=80)

        self.task_tree.heading("Hinweise", text="Hinweise", command=lambda: self._sort_task_tree_by_column("Hinweise"))
        self.task_tree.column("Hinweise", width=260, minwidth=160)
        
        self.task_tree.heading("Suchbegriffe", text="Suchbegriffe", command=lambda: self._sort_task_tree_by_column("Suchbegriffe"))
        self.task_tree.column("Suchbegriffe", width=200, minwidth=140)
        
        v_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        h_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.task_tree.xview)
        self.task_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.task_tree.bind('<Double-1>', self._on_task_double_click)

        self.task_tree.grid(row=1, column=0, sticky="wens")
        v_scrollbar.grid(row=1, column=1, sticky="ns")
        h_scrollbar.grid(row=2, column=0, sticky="we")
        
        # Aktionen (visuell gruppiert)
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=4, column=0, columnspan=2, sticky="we", pady=(10, 0))

        select_group = ttk.LabelFrame(action_frame, text="Auswahl", padding="8")
        select_group.grid(row=0, column=0, sticky="w", padx=(0, 10), pady=(0, 8))
        review_group = ttk.LabelFrame(action_frame, text="Prüfung", padding="8")
        review_group.grid(row=1, column=0, sticky="w", padx=(0, 10))
        export_group = ttk.LabelFrame(action_frame, text="Export", padding="8")
        export_group.grid(row=0, column=1, rowspan=2, sticky="e")

        self._action_groups = [select_group, review_group, export_group]
        for group in self._action_groups:
            group.configure(width=self._action_group_base_width)
            group.grid_propagate(False)

        action_frame.columnconfigure(0, weight=1)
        action_frame.columnconfigure(1, weight=1)

        self.btn_filter = ttk.Button(select_group, text="Aufgaben filtern", command=self.filter_tasks)
        self.btn_filter.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_select_all = ttk.Button(select_group, text="Alle auswählen", command=self.select_all_tasks)
        self.btn_select_all.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_select_blockers = ttk.Button(select_group, text="Nur Blocker auswählen", command=self.select_blocker_tasks)
        self.btn_select_blockers.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_deselect_all = ttk.Button(select_group, text="Auswahl aufheben", command=self.deselect_all_tasks)
        self.btn_deselect_all.pack(side=tk.LEFT)

        self.btn_bulk_edit = ttk.Button(review_group, text="Mehrfach bearbeiten", command=self.bulk_edit_selected_tasks)
        self.btn_bulk_edit.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_preview_selected = ttk.Button(review_group, text="Auswahl prüfen", command=self.preview_selected_task)
        self.btn_preview_selected.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_approve = ttk.Button(review_group, text="Auswahl freigeben", command=self.approve_selected_tasks)
        self.btn_approve.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_clear_approvals = ttk.Button(review_group, text="Freigaben löschen", command=self.clear_task_approvals)
        self.btn_clear_approvals.pack(side=tk.LEFT)

        self.btn_preview_lek = ttk.Button(export_group, text="Gesamtausgabe prüfen", command=self.preview_lek_output)
        self.btn_preview_lek.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_export_selected = ttk.Button(export_group, text="Auswahl exportieren", command=self.export_selected)
        self.btn_export_selected.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_export_all = ttk.Button(export_group, text="Freigaben exportieren", command=self.export_all)
        self.btn_export_all.pack(side=tk.LEFT)

        self._default_button_labels = {
            self.btn_browse: "Datei wählen & laden...",
            self.btn_filter: "Aufgaben filtern",
            self.btn_approve: "Auswahl freigeben",
            self.btn_export_all: "Freigaben exportieren",
        }
        
        # Grid-Konfiguration für responsives Layout
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self._sync_duplicate_mode_ui()
        self._apply_ui_mode()
        self._attach_quick_tooltips()
        self.root.bind('<Configure>', self._on_root_resize, add='+')
        self.root.after_idle(self._apply_adaptive_action_group_width)
        self._set_step(1, show_message=False)

    def _on_root_resize(self, _event=None):
        """Reagiert gedrosselt auf Fenstergrößenänderungen für adaptive Gruppenbreiten."""
        if self._resize_after_id is not None:
            try:
                self.root.after_cancel(self._resize_after_id)
            except Exception:
                pass

        self._resize_after_id = self.root.after(80, self._apply_adaptive_action_group_width)

    def _apply_adaptive_action_group_width(self):
        """Passt die Breite der Aktionsgruppen an die aktuelle Fensterbreite an."""
        self._resize_after_id = None

        groups = list(getattr(self, '_action_groups', []) or [])
        if not groups:
            return

        try:
            window_width = int(self.root.winfo_width() or 0)
        except Exception:
            window_width = 0

        if window_width <= 0:
            return

        # Linke/rechte Außenabstände + mittlerer Spaltenabstand grob berücksichtigen
        estimated_side_padding = 80
        estimated_middle_gap = 20
        available_per_column = (window_width - estimated_side_padding - estimated_middle_gap) // 2

        target_width = min(self._action_group_base_width, available_per_column)
        target_width = max(self._action_group_min_width, target_width)

        for group in groups:
            try:
                group.configure(width=target_width)
            except Exception:
                continue

    def _attach_quick_tooltips(self):
        """Hängt kurze Hover-Infotexte an zentrale Bedienbuttons."""
        tooltip_map = {
            self.ui_mode_combo: "Hier wählen Sie den Arbeitsmodus: 'einfach' für den Standardablauf, 'erweitert' für zusätzliche Import- und Diagnosefunktionen.",
            self.entry_file_path: "Zeigt den Pfad zur aktuellen Aufgabensammlung. Sie können hier auch manuell einen .docx-Pfad eintragen.",
            self.btn_browse: "Öffnet den Dateidialog, übernimmt die gewählte Datei und lädt die Aufgaben direkt in den Assistenten.",
            self.btn_import_task: "Importiert genau eine neue Aufgabe aus einer separaten Word-Datei in die aktuell geladene Sammlung (nur erweitert).",
            self.btn_import_tasks_bulk: "Importiert mehrere Word-Dateien nacheinander als Aufgabenserie in die aktuelle Sammlung (nur erweitert).",
            self.btn_authoring_help: "Zeigt eine kurze Eingabehilfe mit Feldregeln, JSON-Referenzen und Formulierungsbeispielen für neue Aufgaben.",
            self.duplicate_mode_combo: "Legt fest, wie streng ähnliche Aufgaben als Duplikat erkannt werden (strict/normal/relaxed).",
            self.btn_apply_duplicate_mode: "Übernimmt das gewählte Duplikat-Preset sofort für diese Sitzung, ohne es dauerhaft zu speichern.",
            self.btn_save_duplicate_mode: "Speichert das aktuelle Duplikat-Preset als Standard, damit es beim nächsten Start automatisch gilt.",
            self.entry_keywords_filter: "Filtern Sie Aufgaben über Schlagworte (kommagetrennt), z. B. VLAN, Routing, Fehlersuche.",
            self.combo_filter_difficulty: "Begrenzt die Anzeige auf einen Schwierigkeitsgrad oder zeigt mit 'Alle' die komplette Trefferliste.",
            self.spin_max_tasks: "Begrenzt die Anzahl der angezeigten Treffer. Nützlich, um Auswahl und Prüfung übersichtlich zu halten.",
            self.chk_only_blockers: "Blendet nur Aufgaben mit blockierenden Warnungen ein, damit Sie kritische Fälle zuerst korrigieren können.",
            self.task_tree: "Zentrale Aufgabenliste: markieren, per Spaltenkopf sortieren und per Doppelklick eine Aufgabe direkt bearbeiten.",
            self.btn_step_prev: "Wechselt zum vorherigen Assistenten-Schritt, ohne den aktuellen Stand zu verlieren.",
            self.btn_step_next: "Wechselt zum nächsten Schritt, sobald die nötigen Voraussetzungen (z. B. Freigaben) erfüllt sind.",
            self.btn_filter: "Wendet die aktuellen Kriterien an und aktualisiert die Aufgabenliste passend zu Ihren Filtereinstellungen.",
            self.btn_select_all: "Markiert alle aktuell sichtbaren Aufgaben, damit Sie sie gemeinsam freigeben oder bearbeiten können.",
            self.btn_select_blockers: "Markiert gezielt nur Aufgaben mit blockierenden Warnungen für eine schnelle Korrekturrunde.",
            self.btn_bulk_edit: "Bearbeitet Kategorie, Schwierigkeit und Schlagworte für mehrere markierte Aufgaben in einem Schritt.",
            self.btn_deselect_all: "Hebt die komplette aktuelle Markierung auf und startet die Auswahl bei Bedarf neu.",
            self.btn_approve: "Gibt die markierten Aufgaben für Vorschau und Export frei und aktualisiert den Wizard-Fortschritt.",
            self.btn_clear_approvals: "Entfernt alle bisherigen Freigaben, wenn Sie den Auswahlprozess neu aufsetzen möchten.",
            self.btn_preview_selected: "Prüft die aktuell ausgewählte Aufgabe als Textvorschau zur schnellen Inhaltskontrolle.",
            self.btn_preview_lek: "Prüft die Gesamtausgabe der freigegebenen Aufgaben in der späteren Exportreihenfolge.",
            self.btn_export_selected: "Exportiert nur die aktuell markierten und freigegebenen Aufgaben in ein LEK-Dokument.",
            self.btn_export_all: "Exportiert alle freigegebenen Aufgaben gesammelt in eine vollständige LEK-Datei.",
        }

        self._tooltips = [HoverTooltip(widget, text) for widget, text in tooltip_map.items()]

    def _is_advanced_mode(self):
        """True, wenn der UI-Modus auf 'erweitert' steht."""
        mode_value = str(self.ui_mode_var.get() or '').strip().lower()
        return mode_value.startswith('erweitert')

    def _on_ui_mode_changed(self, _event=None):
        """Reagiert auf Moduswechsel (einfach/erweitert)."""
        self._apply_ui_mode()

        if not self._is_advanced_mode() and bool(self.only_blockers_var.get()):
            self.only_blockers_var.set(False)
            if self.loaded_tasks:
                self.filter_tasks(show_message=False)

    def _apply_ui_mode(self):
        """Aktualisiert Verfügbarkeit erweiterter Funktionen je nach UI-Modus."""
        advanced_enabled = self._is_advanced_mode()

        if hasattr(self, 'ui_mode_hint_var'):
            if advanced_enabled:
                self.ui_mode_hint_var.set("Erweitert zeigt Import-/Diagnosefunktionen an.")
            else:
                self.ui_mode_hint_var.set("Einfach (empfohlen) für den Standard-Workflow.")

        advanced_grid_widgets = [
            self.btn_import_task,
            self.btn_import_tasks_bulk,
            self.btn_authoring_help,
            self.lbl_duplicate_mode,
            self.duplicate_mode_combo,
            self.btn_apply_duplicate_mode,
            self.btn_save_duplicate_mode,
            self.lbl_duplicate_threshold,
            self.chk_only_blockers,
        ]

        if advanced_enabled:
            for widget in advanced_grid_widgets:
                try:
                    widget.grid()
                except Exception:
                    pass

            self.duplicate_mode_combo.configure(state='readonly')
            self.btn_apply_duplicate_mode.configure(state=tk.NORMAL)
            self.btn_save_duplicate_mode.configure(state=tk.NORMAL)

            if not self.btn_select_blockers.winfo_ismapped():
                self.btn_select_blockers.pack(side=tk.LEFT, padx=(0, 8))
        else:
            for widget in advanced_grid_widgets:
                try:
                    widget.grid_remove()
                except Exception:
                    pass

            self.duplicate_mode_combo.configure(state='disabled')
            self.btn_apply_duplicate_mode.configure(state=tk.DISABLED)
            self.btn_save_duplicate_mode.configure(state=tk.DISABLED)

            if self.btn_select_blockers.winfo_ismapped():
                self.btn_select_blockers.pack_forget()
            self.btn_select_blockers.configure(state=tk.DISABLED)

        self._update_wizard_step_ui()

    def _available_duplicate_modes(self):
        """Liefert verfügbare Duplikat-Presets aus den Regeln."""
        thresholds = self._rule_value('duplicate_similarity_thresholds', {})
        if isinstance(thresholds, dict) and thresholds:
            return [str(mode) for mode in thresholds.keys()]
        return ['strict', 'normal', 'relaxed']

    def _sync_duplicate_mode_ui(self):
        """Synchronisiert Preset-Auswahl und aktive Schwellenwert-Anzeige."""
        modes = self._available_duplicate_modes()
        if hasattr(self, 'duplicate_mode_combo'):
            self.duplicate_mode_combo.configure(values=modes)

        current_mode = str(self._rule_value('duplicate_mode', 'normal') or 'normal').lower()
        if current_mode not in modes and modes:
            current_mode = modes[0]

        if hasattr(self, 'duplicate_mode_var'):
            self.duplicate_mode_var.set(current_mode)

        threshold = '-'
        resolver = getattr(self.word_processor, '_resolve_duplicate_threshold', None)
        if callable(resolver):
            try:
                resolved = resolver()
                threshold = f"{float(str(resolved)):.2f}"
            except Exception:
                threshold = '-'

        if hasattr(self, 'duplicate_threshold_var'):
            self.duplicate_threshold_var.set(
                f"Aktive Duplikat-Schwelle: {threshold} (Modus: {current_mode})"
            )

    def _apply_duplicate_mode_from_ui(self, show_message=True):
        """Übernimmt den im UI gewählten Duplikat-Modus in die Laufzeitregeln."""
        if not self._is_advanced_mode():
            messagebox.showinfo("Hinweis", "Duplikat-Presets sind nur im erweiterten Modus verfügbar.")
            return

        selected = str(self.duplicate_mode_var.get() or '').strip().lower()
        if not selected:
            return

        modes = self._available_duplicate_modes()
        if selected not in modes:
            messagebox.showwarning(
                "Ungültiges Preset",
                f"Der Modus '{selected}' ist nicht verfügbar.",
            )
            return

        if not isinstance(getattr(self.word_processor, 'rules', None), dict):
            self.word_processor.rules = {}
        self.word_processor.rules['duplicate_mode'] = selected

        self._sync_duplicate_mode_ui()
        if show_message:
            messagebox.showinfo(
                "Preset aktiv",
                f"Duplikat-Preset '{selected}' wurde für diese Sitzung aktiviert.",
            )

    def _save_duplicate_mode_as_default(self, show_message=True):
        """Speichert den aktuell gewählten Duplikat-Modus dauerhaft in der Konfiguration."""
        if not self._is_advanced_mode():
            messagebox.showinfo("Hinweis", "Duplikat-Presets sind nur im erweiterten Modus verfügbar.")
            return False

        # Sicherstellen, dass der UI-Wert in die Laufzeitregeln übernommen wurde
        self._apply_duplicate_mode_from_ui(show_message=False)

        persister = getattr(self.word_processor, 'persist_import_rules', None)
        if not callable(persister):
            messagebox.showerror(
                "Speichern fehlgeschlagen",
                "Persistenzfunktion ist nicht verfügbar.",
            )
            return False

        try:
            config_path = persister(runtime_base=_runtime_base_dir())
            self._sync_duplicate_mode_ui()
            if show_message:
                messagebox.showinfo(
                    "Standard gespeichert",
                    "Duplikat-Preset wurde dauerhaft gespeichert.\n\n"
                    f"Datei: {config_path}",
                )
            return True
        except Exception as e:
            messagebox.showerror(
                "Speichern fehlgeschlagen",
                f"Konfiguration konnte nicht gespeichert werden:\n{str(e)}",
            )
            return False

    def _max_reachable_step(self):
        """Ermittelt den maximal zulässigen Wizard-Schritt anhand des aktuellen Zustands."""
        if not self.import_session:
            return 1

        stats = self.import_session.get_stats()
        total = stats.get('total', 0)
        approved = stats.get('approved', 0)

        if total <= 0:
            return 1
        if approved <= 0:
            return 3
        return 4

    def _set_step(self, target_step, show_message=True):
        """Setzt den aktiven Wizard-Schritt, sofern die Voraussetzungen erfüllt sind."""
        target_step = max(1, min(4, int(target_step)))
        max_step = self._max_reachable_step()
        if target_step > max_step:
            if show_message:
                hints = {
                    2: "Bitte zuerst eine gültige Aufgabensammlung laden.",
                    3: "Bitte zuerst die Erkennung prüfen und Aufgaben auswählen.",
                    4: "Bitte zuerst Aufgaben freigeben.",
                }
                messagebox.showwarning("Schritt nicht verfügbar", hints.get(target_step, "Schritt aktuell nicht verfügbar."))
            return False

        self.current_step = target_step
        self._update_wizard_step_ui()
        return True

    def _on_prev_step(self):
        """Navigiert einen Schritt zurück."""
        self._set_step(self.current_step - 1, show_message=False)

    def _on_next_step(self):
        """Navigiert einen Schritt vorwärts, sofern erlaubt."""
        if self.current_step >= 4:
            return
        self._set_step(self.current_step + 1, show_message=True)

    def _next_step_block_reason(self):
        """Liefert einen kurzen, verständlichen Sperrgrund für den nächsten Schritt."""
        if not self.import_session:
            return "Bitte zuerst eine Aufgabensammlung laden."

        stats = self.import_session.get_stats()
        total = int(stats.get('total', 0) or 0)
        approved = int(stats.get('approved', 0) or 0)

        if self.current_step <= 1 and total <= 0:
            return "Bitte zuerst eine Datei laden und Aufgaben erkennen lassen."

        if self.current_step == 2 and total <= 0:
            return "Es wurden noch keine verwertbaren Aufgaben erkannt."

        if self.current_step == 3 and approved <= 0:
            return "Bitte mindestens eine Aufgabe freigeben."

        return ""

    def _update_wizard_step_ui(self):
        """Aktualisiert Schrittanzeige, Navigations- und Aktionsbuttons."""
        label = self.step_labels.get(self.current_step, "Unbekannt")
        step_progress_pct = int(round((self.current_step / 4) * 100))
        self.wizard_step_var.set(f"Schritt {self.current_step}/4 • {step_progress_pct}%: {label}")
        advanced_mode = self._is_advanced_mode()

        max_step = self._max_reachable_step()

        self.btn_step_prev.configure(state=tk.NORMAL if self.current_step > 1 else tk.DISABLED)
        next_enabled = self.current_step < 4 and self.current_step < max_step
        self.btn_step_next.configure(state=tk.NORMAL if next_enabled else tk.DISABLED)

        has_session = self.import_session is not None and max_step >= 2
        enable_workflow = tk.NORMAL if has_session and self.current_step >= 2 else tk.DISABLED

        self.btn_filter.configure(state=enable_workflow)
        self.btn_select_all.configure(state=enable_workflow)
        blockers_state = tk.NORMAL if (enable_workflow == tk.NORMAL and self._is_advanced_mode()) else tk.DISABLED
        self.btn_select_blockers.configure(state=blockers_state)
        self.btn_bulk_edit.configure(state=enable_workflow)
        self.btn_deselect_all.configure(state=enable_workflow)
        self.btn_approve.configure(state=enable_workflow)
        self.btn_clear_approvals.configure(state=enable_workflow)
        self.btn_preview_selected.configure(state=enable_workflow)

        approved_count = 0
        if self.import_session is not None:
            try:
                approved_count = int(self.import_session.get_stats().get('approved', 0))
            except Exception:
                approved_count = 0

        preview_lek_enabled = tk.NORMAL if has_session and self.current_step >= 3 and approved_count > 0 else tk.DISABLED
        self.btn_preview_lek.configure(state=preview_lek_enabled)

        export_enabled = tk.NORMAL if has_session and self.current_step >= 4 and max_step >= 4 else tk.DISABLED
        self.btn_export_selected.configure(state=export_enabled)
        self.btn_export_all.configure(state=export_enabled)

        if self.current_step <= 1:
            self.wizard_hint_var.set("Nächster Schritt: Aufgabensammlung auswählen und laden.")
            self._mark_primary_action(self.btn_browse)
        elif self.current_step == 2:
            step2_hint = (
                "Nächster Schritt: Aufgaben prüfen/filtern, optional Blocker fokussieren und gewünschte Aufgaben markieren."
                if advanced_mode
                else "Nächster Schritt: Aufgaben prüfen/filtern und gewünschte Aufgaben markieren."
            )
            reason = self._next_step_block_reason()
            if reason:
                self.wizard_hint_var.set(
                    f"{step2_hint} "
                    f"(Weiter gesperrt: {reason})"
                )
            else:
                self.wizard_hint_var.set(step2_hint)
            self._mark_primary_action(self.btn_filter)
        elif self.current_step == 3:
            if approved_count > 0:
                if advanced_mode:
                    self.wizard_hint_var.set(
                        "Nächster Schritt: Freigaben prüfen, optional 'Gesamtausgabe prüfen' öffnen und mit 'Weiter' zum Export wechseln."
                    )
                else:
                    self.wizard_hint_var.set("Nächster Schritt: Freigaben prüfen und mit 'Weiter' zum Export wechseln.")
            else:
                reason = self._next_step_block_reason() or "Bitte mindestens eine Aufgabe freigeben."
                self.wizard_hint_var.set(
                    "Nächster Schritt: Markierte Aufgaben über 'Auswahl freigeben' bestätigen. "
                    f"(Weiter gesperrt: {reason})"
                )
            self._mark_primary_action(self.btn_approve)
        else:
            if advanced_mode:
                self.wizard_hint_var.set(
                    "Nächster Schritt: Export mit 'Freigaben exportieren' oder 'Auswahl exportieren' starten (optional vorher Gesamtausgabe prüfen)."
                )
            else:
                self.wizard_hint_var.set("Nächster Schritt: Export mit 'Freigaben exportieren' oder 'Auswahl exportieren' starten.")
            self._mark_primary_action(self.btn_export_all)

    def _mark_primary_action(self, primary_button):
        """Hebt die Primäraktion im aktuellen Wizard-Schritt sichtbar hervor."""
        for button, label in self._default_button_labels.items():
            try:
                button.configure(text=label)
            except Exception:
                pass

        if primary_button in self._default_button_labels:
            try:
                base_label = self._default_button_labels[primary_button]
                primary_button.configure(text=f"➡ {base_label}")
            except Exception:
                pass

    def _build_wizard_status_line(self, approved, total, warnings_total, blocking_total, todo_hint):
        """Erzeugt eine kompakte, visuell priorisierte Wizard-Statuszeile."""
        try:
            approved_num = int(approved or 0)
        except Exception:
            approved_num = 0

        try:
            total_num = int(total or 0)
        except Exception:
            total_num = 0

        try:
            warnings_num = int(warnings_total or 0)
        except Exception:
            warnings_num = 0

        try:
            blocking_num = int(blocking_total or 0)
        except Exception:
            blocking_num = 0

        if total_num > 0:
            progress_pct = int(round((approved_num / total_num) * 100))
        else:
            progress_pct = 0

        if blocking_num > 0:
            state_icon = "⛔"
        elif warnings_num > 0:
            state_icon = "⚠"
        else:
            state_icon = "✅"

        return (
            f"{state_icon} Wizard-Status: Fortschritt {approved_num}/{total_num} ({progress_pct}%)"
            f" | ⚠ Warnungen {warnings_num} | ⛔ Blocker {blocking_num}{todo_hint}"
        )

    def _refresh_wizard_status(self):
        """Aktualisiert die kompakte Statuszeile für den Wizard-Zustand."""
        if not self.import_session:
            self.wizard_status_var.set("ℹ Wizard-Status: Keine Datei geladen")
            self._update_wizard_step_ui()
            return

        stats = self.import_session.get_stats()
        warnings_total = stats.get('warnings', 0)
        blocking_total = stats.get('blocking', 0)
        approved = stats.get('approved', 0)
        total = stats.get('total', 0)
        todo_hint = self._format_blocking_todo_hint()

        self.wizard_status_var.set(
            self._build_wizard_status_line(approved, total, warnings_total, blocking_total, todo_hint)
        )
        self._update_wizard_step_ui()

    def _warning_priority(self, warning_text):
        """Liefert Sortierpriorität für Warnungen (kleiner = wichtiger)."""
        text = str(warning_text or '').strip().lower()
        if not text:
            return 99

        if 'pflichtfeld fehlt' in text:
            return 0
        if 'kategorie fehlt' in text:
            return 1
        if 'doppelte aufgabennummer' in text:
            return 2
        if 'inkonsistenter schwierigkeitsgrad' in text:
            return 3
        if 'keinen verwertbaren inhalt' in text:
            return 4
        if 'einleitungs-/kontexttext' in text:
            return 6
        return 10

    def _format_warning_preview(self, warnings, hints=None, max_items=2):
        """Erzeugt eine kompakte Vorschau für Warnungen und Hinweise in der Tabelle."""
        normalized_warnings = [str(w).strip() for w in (warnings or []) if str(w).strip()]
        normalized_hints = [str(h).strip() for h in (hints or []) if str(h).strip()]

        parts = []
        if normalized_warnings:
            sorted_warnings = sorted(normalized_warnings, key=lambda w: (self._warning_priority(w), w.lower()))
            top = sorted_warnings[:max_items]
            preview = '; '.join(top)
            remaining = len(sorted_warnings) - len(top)
            if remaining > 0:
                preview += f"; +{remaining} weitere"
            parts.append(preview)

        if normalized_hints:
            # Hinweise sichtbar kennzeichnen und kompakt halten
            hint_text = '; '.join(f"Hinweis: {text}" for text in normalized_hints[:max_items])
            hint_remaining = len(normalized_hints) - min(len(normalized_hints), max_items)
            if hint_remaining > 0:
                hint_text += f"; +{hint_remaining} weitere Hinweise"
            parts.append(hint_text)

        return ' | '.join(parts)

    def _format_blocking_todo_hint(self):
        """Formatiert eine knappe To-fix-Zusammenfassung für blockierende Warnungen."""
        if not self.import_session:
            return ""

        summary_getter = getattr(self.import_session, 'get_blocking_summary', None)
        if not callable(summary_getter):
            return ""

        raw_summary = summary_getter()
        if not isinstance(raw_summary, dict):
            return ""

        summary: dict[str, Any] = raw_summary
        parts = []

        title_count = int(summary.get('missing_title', 0) or 0)
        if title_count > 0:
            parts.append(f"Titel {title_count}")

        category_count = int(summary.get('missing_category', 0) or 0)
        if category_count > 0:
            parts.append(f"Kategorie {category_count}")

        diff_count = int(summary.get('inconsistent_difficulty', 0) or 0)
        if diff_count > 0:
            parts.append(f"Schwierigkeit {diff_count}")

        required_count = int(summary.get('missing_required', 0) or 0)
        if required_count > 0:
            parts.append(f"Pflichtfelder {required_count}")

        content_count = int(summary.get('missing_content', 0) or 0)
        if content_count > 0:
            parts.append(f"Inhalt {content_count}")

        duplicate_numbering_count = int(summary.get('duplicate_numbering', 0) or 0)
        if duplicate_numbering_count > 0:
            parts.append(f"Nummerierung {duplicate_numbering_count}")

        if not parts:
            return ""

        return " | To-fix: " + ", ".join(parts)

    def _task_has_blocking_warning(self, task):
        """Prüft, ob eine Aufgabe mindestens eine blockierende Warnung enthält."""
        warnings = list(task.get('warnings', []) or [])
        if not warnings:
            warnings = list(task.get('pre_warnings', []) or [])

        return any(is_blocking_warning(w) for w in warnings)

    def _on_toggle_only_blockers(self):
        """Aktualisiert die Ansicht beim Umschalten des Blocker-Filters."""
        if not self._is_advanced_mode():
            self.only_blockers_var.set(False)
            return

        if not self.loaded_tasks:
            return
        self.filter_tasks()

    def _default_task_authoring_hints(self):
        """Liefert Standard-Hinweise für intuitive Aufgabenerstellung."""
        return {
            "json_files": [
                {
                    "path": "data/config/import_rules.json",
                    "purpose": "Regelt Pflichtfelder, Alias-Namen, Schwierigkeitswerte, Vorschau-/Export-Reihenfolge sowie Export-Layout.",
                    "sections": [
                        "template_rules.required_fields",
                        "difficulty_rules.allowed_values",
                        "field_alias_rules.structured_task_fields",
                        "preview_rules.task_flow_sections",
                        "export_rules.title_points_box.min_inner_width",
                        "export_rules.title_points_box.padding_spaces",
                    ],
                },
                {
                    "path": "data/config/task_authoring_hints.json",
                    "purpose": "Projektweite Formulierungs- und Eingabehilfe für neue Aufgaben.",
                    "sections": [
                        "formulation_tips",
                        "task_description_template",
                        "title_examples",
                    ],
                },
            ],
            "formulation_tips": [
                "Titel kurz und eindeutig (max. ~8 Wörter, ohne Floskeln).",
                "Aufgabenstellung mit klarer Handlungsanweisung beginnen (z. B. 'Analysieren Sie ...').",
                "Einleitung/Kontext nur ergänzen, wenn sie für die Lösung wirklich nötig ist.",
                "Hinweis nur als Hilfe formulieren, nicht als Vorwegnahme der Lösung.",
                "Schlagworte konkret und kommagetrennt pflegen (z. B. VLAN, Routing, Fehlersuche).",
            ],
            "task_description_template": [
                "Ausgangslage: <kurzer Kontext>",
                "Aufgabe: <konkrete Arbeitsanweisung>",
                "Ergebnisformat: <z. B. Tabelle, Stichpunkte, Begründung>",
                "Bewertungskriterium: <was wird als korrekt gewertet?>",
            ],
            "title_examples": [
                "Subnetz planen für zwei Standorte",
                "Fehleranalyse bei DHCP-Ausfall",
                "Sicheres Passwortkonzept begründen",
            ],
            "field_hints": {
                "category": "Beispiel: Netzwerktechnik, IT-Sicherheit, Office-Praxis",
                "difficulty": "Erlaubte Werte: leicht | mittel | schwer",
                "keywords": "Beispiel: VLAN, Routing, Fehlersuche",
                "title": "Kurz und konkret, z. B. 'Subnetzplanung Filiale Nord'",
            },
            "category_defaults": {
                "netzwerktechnik": {
                    "keywords": ["VLAN", "Routing", "Fehlersuche"],
                    "difficulty": "mittel",
                    "title_prefix": "Netzwerkaufgabe:",
                },
                "it-sicherheit": {
                    "keywords": ["Authentifizierung", "Richtlinie", "Risikoanalyse"],
                    "difficulty": "mittel",
                    "title_prefix": "Security-Aufgabe:",
                },
                "office-praxis": {
                    "keywords": ["Excel", "Word", "Dokumentation"],
                    "difficulty": "leicht",
                    "title_prefix": "Office-Aufgabe:",
                },
            },
        }

    def _load_task_authoring_hints(self):
        """Lädt optionale Authoring-Hinweise aus JSON mit robustem Fallback."""
        defaults = self._default_task_authoring_hints()
        hints_path = _runtime_base_dir() / "data" / "config" / "task_authoring_hints.json"

        if not hints_path.exists():
            return defaults

        try:
            with hints_path.open('r', encoding='utf-8') as f:
                loaded = json.load(f)
            if not isinstance(loaded, dict):
                return defaults

            merged = dict(defaults)
            for key in ('json_files', 'formulation_tips', 'task_description_template', 'title_examples'):
                value = loaded.get(key)
                if isinstance(value, list) and value:
                    merged[key] = value

            field_hints = loaded.get('field_hints')
            if isinstance(field_hints, dict) and field_hints:
                normalized_hints = {
                    str(k).strip().lower(): str(v).strip()
                    for k, v in field_hints.items()
                    if str(k).strip() and str(v).strip()
                }
                if normalized_hints:
                    merged['field_hints'] = normalized_hints

            category_defaults = loaded.get('category_defaults')
            if isinstance(category_defaults, dict) and category_defaults:
                normalized_defaults = {}
                for category_key, cfg in category_defaults.items():
                    key = str(category_key).strip().lower()
                    if not key or not isinstance(cfg, dict):
                        continue
                    normalized_defaults[key] = dict(cfg)
                if normalized_defaults:
                    merged['category_defaults'] = normalized_defaults
            return merged
        except Exception:
            return defaults

    def _user_config_path(self):
        """Liefert den Pfad zur benutzerspezifischen Authoring-Config."""
        return _runtime_base_dir() / "data" / "config" / "task_authoring_user_config.json"

    def _load_authoring_user_config(self):
        """Lädt die benutzerspezifische Authoring-Config robust mit Fallback."""
        cfg_path = self._user_config_path()
        if not cfg_path.exists():
            return {}

        try:
            with cfg_path.open('r', encoding='utf-8') as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _save_authoring_user_config(self, config):
        """Speichert die benutzerspezifische Authoring-Config."""
        if not isinstance(config, dict):
            return False

        cfg_path = self._user_config_path()
        try:
            cfg_path.parent.mkdir(parents=True, exist_ok=True)
            with cfg_path.open('w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                f.write('\n')
            return True
        except Exception:
            return False

    def _get_last_used_category(self):
        """Liest die zuletzt verwendete Kategorie aus der User-Config."""
        cfg = self._load_authoring_user_config()
        return str(cfg.get('last_used_category', '') or '').strip()

    def _set_last_used_category(self, category):
        """Persistiert die zuletzt verwendete Kategorie in der User-Config."""
        value = str(category or '').strip()
        if not value:
            return False

        cfg = self._load_authoring_user_config()
        cfg['last_used_category'] = value
        return self._save_authoring_user_config(cfg)

    def show_task_authoring_help(self):
        """Zeigt eine kompakte Eingabehilfe für neue Aufgaben inklusive JSON-Hinweisen."""
        hints = self._load_task_authoring_hints()

        json_files = hints.get('json_files', []) or []
        formulation_tips = hints.get('formulation_tips', []) or []
        template_lines = hints.get('task_description_template', []) or []
        title_examples = hints.get('title_examples', []) or []
        field_hints = hints.get('field_hints', {}) or {}

        lines = [
            "Eingabehilfe: Neue Aufgabe intuitiv anlegen",
            "",
            "Geeignete JSON-Dateien:",
        ]

        if json_files:
            for entry in json_files:
                if not isinstance(entry, dict):
                    continue
                path = str(entry.get('path', '')).strip()
                purpose = str(entry.get('purpose', '')).strip()
                sections = entry.get('sections', []) or []
                section_text = ", ".join(str(s).strip() for s in sections if str(s).strip())

                if path:
                    lines.append(f"- {path}")
                if purpose:
                    lines.append(f"  Zweck: {purpose}")
                if section_text:
                    lines.append(f"  Relevante Bereiche: {section_text}")
        else:
            lines.append("- data/config/import_rules.json")

        lines.extend(["", "Formulierungs-Tipps:"])
        if formulation_tips:
            lines.extend(f"- {tip}" for tip in formulation_tips)
        else:
            lines.append("- Aufgaben klar, präzise und handlungsorientiert formulieren.")

        if template_lines:
            lines.extend(["", "Vorschlag für Aufgabenbeschreibung:"])
            lines.extend(f"- {line}" for line in template_lines)

        if title_examples:
            lines.extend(["", "Titelbeispiele:"])
            lines.extend(f"- {example}" for example in title_examples)

        if isinstance(field_hints, dict) and field_hints:
            lines.extend(["", "Feldhilfen (Kurz):"])
            hint_order = [
                ('category', 'Kategorie'),
                ('difficulty', 'Schwierigkeit'),
                ('keywords', 'Schlagworte'),
                ('title', 'Titel'),
            ]
            for key, label in hint_order:
                hint_text = str(field_hints.get(key, '') or '').strip()
                if hint_text:
                    lines.append(f"- {label}: {hint_text}")

        messagebox.showinfo("Eingabehilfe", "\n".join(lines))
    
    def browse_and_load_file(self):
        """Öffnet einen Dateidialog zur Auswahl der Word-Datei und lädt sie direkt"""
        # Aufgaben-Verzeichnis als Standard setzen
        base_dir = _runtime_base_dir()
        aufgaben_dir = base_dir / "data" / "Aufgaben"
        if not aufgaben_dir.exists():
            aufgaben_dir = base_dir  # Fallback auf Hauptverzeichnis
        
        filename = filedialog.askopenfilename(
            title="Word-Datei auswählen",
            initialdir=str(aufgaben_dir),
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
            self.load_tasks()
    
    def browse_file(self):
        """Öffnet einen Dateidialog zur Auswahl der Word-Datei"""
        # Aufgaben-Verzeichnis als Standard setzen
        base_dir = _runtime_base_dir()
        aufgaben_dir = base_dir / "data" / "Aufgaben"
        if not aufgaben_dir.exists():
            aufgaben_dir = base_dir  # Fallback auf Hauptverzeichnis
        
        filename = filedialog.askopenfilename(
            title="Word-Datei auswählen",
            initialdir=str(aufgaben_dir),
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
    
    def load_tasks(self, allow_title_migration_prompt=True):
        """Lädt Aufgaben aus der ausgewählten Word-Datei"""
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("Fehler", "Bitte wählen Sie eine Word-Datei aus.")
            return
        
        if not os.path.exists(file_path):
            messagebox.showerror("Fehler", "Die ausgewählte Datei existiert nicht.")
            return
        
        try:
            diagnostics_report = None
            self.loaded_tasks = []

            if hasattr(self.word_processor, 'extract_tasks_with_diagnostics'):
                self.loaded_tasks, diagnostics_report = self.word_processor.extract_tasks_with_diagnostics(file_path)
            else:
                self.loaded_tasks = self.word_processor.extract_tasks(file_path)

            self.current_displayed_tasks = self.loaded_tasks  # Setze initial alle Aufgaben als angezeigt

            missing_title_tasks = []
            for task in self.loaded_tasks:
                warnings = task.get('warnings', []) or task.get('pre_warnings', []) or []
                if any('Pflichtfeld fehlt: Titel' in str(w) for w in warnings):
                    missing_title_tasks.append(task)

            if allow_title_migration_prompt and missing_title_tasks and hasattr(self.word_processor, 'migrate_missing_titles_in_collection'):
                auto_fix = messagebox.askyesno(
                    "Titel automatisch ergänzen",
                    f"Es wurden {len(missing_title_tasks)} Aufgabe(n) ohne Titel erkannt.\n\n"
                    "Soll der Titel automatisch aus der Aufgabenstellung ergänzt werden?\n"
                    "Vor der Änderung wird ein Backup erstellt.",
                )
                if auto_fix:
                    try:
                        result = self.word_processor.migrate_missing_titles_in_collection(file_path)
                        if result.get('changed_tasks', 0) > 0:
                            updated_ids = result.get('updated_ids', []) or []
                            updated_entries = result.get('updated_entries', []) or []
                            ids_preview = ''
                            if updated_ids:
                                preview_list = updated_ids[:10]
                                ids_preview = "\n\nBetroffene Aufgaben-IDs:\n- " + "\n- ".join(preview_list)
                                if len(updated_ids) > 10:
                                    ids_preview += f"\n- ... (+{len(updated_ids) - 10} weitere)"

                            entry_preview = ''
                            if updated_entries:
                                pairs = []
                                for item in updated_entries[:5]:
                                    item_id = str(item.get('id', '') or '-').strip() or '-'
                                    item_title = str(item.get('title', '') or 'Ohne Titel').strip() or 'Ohne Titel'
                                    pairs.append(f"- {item_id} → {item_title}")
                                entry_preview = "\n\nAbgeleitete Titel (Auszug):\n" + "\n".join(pairs)
                                if len(updated_entries) > 5:
                                    entry_preview += f"\n- ... (+{len(updated_entries) - 5} weitere)"

                            messagebox.showinfo(
                                "Titel ergänzt",
                                f"Titel wurden für {result.get('changed_tasks', 0)} Aufgabe(n) ergänzt.\n"
                                f"Backup: {result.get('backup_file', '-')}"
                                f"{ids_preview}"
                                f"{entry_preview}",
                            )
                            self.load_tasks(allow_title_migration_prompt=False)
                            return
                    except Exception as migration_error:
                        messagebox.showerror(
                            "Migration fehlgeschlagen",
                            f"Automatische Titelergänzung fehlgeschlagen:\n{str(migration_error)}",
                        )
            
            # Extrahiere LEK-Thema aus dem Dateinamen
            self.source_filename = os.path.basename(file_path)
            self.lek_theme = self._extract_lek_theme(self.source_filename)

            # Wizard-Session vorbereiten (für geführten Import in Sprint 1+)
            self.import_session = ImportSession.from_raw_tasks(
                source_file=file_path,
                source_filename=self.source_filename,
                lek_theme=self.lek_theme,
                raw_tasks=self.loaded_tasks,
                global_warnings=(diagnostics_report or {}).get('global_warnings', []),
            )
            
            self.populate_task_tree(self.loaded_tasks)
            self._refresh_wizard_status()
            stats = self.import_session.get_stats() if self.import_session else {"total": len(self.loaded_tasks)}

            warning_hint = ""
            if diagnostics_report:
                warning_hint = (
                    f"\nWarnungsbehaftete Aufgaben: {diagnostics_report.get('warning_task_count', 0)}"
                )

            if stats.get('total', 0) == 0:
                self._set_step(1, show_message=False)
                messagebox.showwarning(
                    "Keine Aufgaben erkannt",
                    "In der gewählten Datei wurden keine Aufgaben erkannt.\n\n"
                    "Bitte prüfen Sie die Struktur:\n"
                    "- Kategorie als 'Überschrift 1'\n"
                    "- Aufgabe als 'Überschrift 2'\n\n"
                    "Empfehlung: Nutzen Sie die Musterdatei\n"
                    "data/Vorlagen/AUFGABEN-Vorlage.docx",
                )
                return

            self._set_step(2, show_message=False)

            duplicate_numbering = (diagnostics_report or {}).get('duplicate_numbering', []) or []
            if duplicate_numbering:
                details = []
                for entry in duplicate_numbering[:8]:
                    label = str(entry.get('number_display') or '').strip() or '-'
                    titles = [str(title or '').strip() for title in (entry.get('titles') or []) if str(title or '').strip()]
                    title_preview = ' | '.join(titles[:3]) if titles else 'Ohne Titel'
                    if len(titles) > 3:
                        title_preview += f" | +{len(titles) - 3} weitere"
                    details.append(f"- {label}: {title_preview}")
                if len(duplicate_numbering) > 8:
                    details.append(f"- ... (+{len(duplicate_numbering) - 8} weitere Dubletten)")

                messagebox.showwarning(
                    "Doppelte Aufgabennummern erkannt",
                    f"{stats['total']} Aufgaben geladen.{warning_hint}\n\n"
                    "Es wurden doppelte Aufgabennummern erkannt.\n"
                    "Bitte passen Sie die Nummerierungen in der Quelldatei an und laden Sie die Sammlung danach erneut.\n\n"
                    "Betroffene Nummern:\n"
                    + "\n".join(details),
                )
                return

            messagebox.showinfo("Erfolg", f"{stats['total']} Aufgaben geladen.{warning_hint}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Datei: {str(e)}")

    def import_task_from_word(self):
        """Übernimmt eine neue Aufgabe aus einer separaten Word-Datei in die aktuelle Sammlung."""
        if not self._is_advanced_mode():
            messagebox.showinfo("Hinweis", "Diese Funktion ist nur im erweiterten Modus verfügbar.")
            return

        target_collection = self._get_valid_import_target_collection()
        if not target_collection:
            return

        source_file = filedialog.askopenfilename(
            title="Quell-Datei mit neuer Aufgabe auswählen",
            initialdir=str(_runtime_base_dir()),
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")],
        )
        if not source_file:
            return

        if os.path.abspath(source_file) == os.path.abspath(target_collection):
            messagebox.showwarning(
                "Hinweis",
                "Quell- und Zieldatei dürfen nicht identisch sein.",
            )
            return

        metadata = self._ask_import_metadata()
        if metadata is None:
            return

        should_import = self._preview_and_confirm_import(
            source_file=source_file,
            target_collection=target_collection,
            category=metadata['category'],
            difficulty=metadata['difficulty'],
            keywords=metadata['keywords'],
            title=metadata.get('title', ''),
            allow_cancel=False,
        )
        if not should_import:
            return

        try:
            result = self.word_processor.append_task_from_document(
                source_doc_path=source_file,
                target_collection_path=target_collection,
                category=metadata['category'],
                difficulty=metadata['difficulty'],
                keywords=metadata['keywords'],
                title=metadata.get('title', ''),
            )
            self.load_tasks()
            messagebox.showinfo(
                "Erfolg",
                "Neue Aufgabe wurde in die Aufgabensammlung übernommen.\n\n"
                f"ID: {result.get('id', '-')}\n"
                f"Titel: {result.get('title', '-')}\n"
                f"Backup: {result.get('backup_file', '-')}",
            )
        except Exception as e:
            messagebox.showerror("Fehler", f"Übernahme fehlgeschlagen: {str(e)}")

    def import_tasks_bulk_from_word(self):
        """Übernimmt mehrere neue Aufgaben aus separaten Word-Dateien in die aktuelle Sammlung."""
        if not self._is_advanced_mode():
            messagebox.showinfo("Hinweis", "Diese Funktion ist nur im erweiterten Modus verfügbar.")
            return

        target_collection = self._get_valid_import_target_collection()
        if not target_collection:
            return

        source_files = filedialog.askopenfilenames(
            title="Quell-Dateien mit neuen Aufgaben auswählen",
            initialdir=str(_runtime_base_dir()),
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")],
        )
        if not source_files:
            return

        source_files = [str(s) for s in source_files]
        source_files = [s for s in source_files if os.path.abspath(s) != os.path.abspath(target_collection)]
        if not source_files:
            messagebox.showwarning("Hinweis", "Keine gültigen Quelldateien ausgewählt.")
            return

        metadata = self._ask_import_metadata()
        if metadata is None:
            return

        imported = 0
        skipped = 0
        failed = []
        overridden = 0
        imported_entries = []
        raw_max_errors: Any = self._rule_value('bulk_max_errors', 5)
        try:
            max_errors = int(raw_max_errors or 5)
        except (TypeError, ValueError):
            max_errors = 5
        bulk_stopped_reason = ""

        for source_file in source_files:
            metadata_for_file = dict(metadata)

            override_answer = messagebox.askyesnocancel(
                "Metadaten pro Datei",
                f"Quelle: {os.path.basename(source_file)}\n\n"
                "Metadaten für diese Datei überschreiben?\n\n"
                "Ja = Metadaten für diese Datei anpassen\n"
                "Nein = Serien-Standard übernehmen\n"
                "Abbrechen = Bulk-Serie stoppen",
            )

            if override_answer is None:
                break

            if override_answer is True:
                custom_metadata = None
                metadata_action = 'apply'
                while True:
                    custom_metadata = self._ask_import_metadata(
                        default_category=metadata['category'],
                        default_difficulty=metadata['difficulty'],
                        default_keywords=metadata['keywords'],
                        default_title=metadata.get('title', ''),
                    )

                    if custom_metadata is not None:
                        break

                    correction_choice = messagebox.askyesnocancel(
                        "Metadaten-Eingabe abgebrochen",
                        f"Quelle: {os.path.basename(source_file)}\n\n"
                        "Möchten Sie die Metadaten erneut bearbeiten?\n\n"
                        "Ja = erneut bearbeiten\n"
                        "Nein = Datei überspringen\n"
                        "Abbrechen = Serie stoppen",
                    )
                    if correction_choice is True:
                        continue
                    if correction_choice is False:
                        skipped += 1
                        metadata_action = 'skip-file'
                        break
                    # Abbrechen => gesamte Serie stoppen
                    metadata_action = 'stop-series'
                    break

                if metadata_action == 'skip-file':
                    continue
                if metadata_action == 'stop-series' or custom_metadata is None:
                    break

                metadata_for_file = custom_metadata
                overridden += 1

            decision = self._preview_and_confirm_import(
                source_file=source_file,
                target_collection=target_collection,
                category=metadata_for_file['category'],
                difficulty=metadata_for_file['difficulty'],
                keywords=metadata_for_file['keywords'],
                title=metadata_for_file.get('title', ''),
                allow_cancel=True,
            )

            # None => Abbruch der gesamten Serie
            if decision is None:
                break
            if decision is False:
                skipped += 1
                continue

            try:
                result = self.word_processor.append_task_from_document(
                    source_doc_path=source_file,
                    target_collection_path=target_collection,
                    category=metadata_for_file['category'],
                    difficulty=metadata_for_file['difficulty'],
                    keywords=metadata_for_file['keywords'],
                    title=metadata_for_file.get('title', ''),
                )
                imported += 1
                imported_entries.append(
                    f"{os.path.basename(source_file)} → {result.get('id', '-') } ({result.get('title', '-')})"
                )
            except Exception as e:
                failed.append(f"{os.path.basename(source_file)}: {str(e)}")
                if max_errors > 0 and len(failed) >= max_errors:
                    bulk_stopped_reason = f"Fehlerlimit erreicht ({len(failed)}/{max_errors})."
                    break

        self.load_tasks()

        details = [
            f"Übernommen: {imported}",
            f"Übersprungen: {skipped}",
            f"Fehler: {len(failed)}",
            f"Metadaten pro Datei angepasst: {overridden}",
        ]
        if failed:
            details.append("\nFehlerdetails:")
            details.extend(f"- {entry}" for entry in failed[:8])
            if len(failed) > 8:
                details.append(f"- ... (+{len(failed) - 8} weitere)")

        if imported_entries:
            details.append("\nImportierte Aufgaben:")
            details.extend(f"- {entry}" for entry in imported_entries[:10])
            if len(imported_entries) > 10:
                details.append(f"- ... (+{len(imported_entries) - 10} weitere)")

        if bulk_stopped_reason:
            details.append(f"\nSerie gestoppt: {bulk_stopped_reason}")

        messagebox.showinfo("Bulk-Übernahme abgeschlossen", "\n".join(details))

    def _get_valid_import_target_collection(self):
        """Liefert den validierten Pfad der aktuell geladenen Ziel-Aufgabensammlung."""
        target_collection = self.file_path_var.get().strip()
        if not target_collection:
            messagebox.showwarning(
                "Hinweis",
                "Bitte zuerst die Ziel-Aufgabensammlung auswählen und laden.",
            )
            return None

        if not os.path.exists(target_collection):
            messagebox.showerror("Fehler", "Die ausgewählte Ziel-Aufgabensammlung existiert nicht.")
            return None

        return target_collection

    def _category_suggestions(self, preferred_category=None, category_defaults=None):
        """Liefert Kategorievorschläge mit Ranking nach Häufigkeit (absteigend)."""
        stats = {}

        def _add(value, weight=1):
            text = str(value or '').strip()
            if not text:
                return
            key = text.lower()
            if key not in stats:
                stats[key] = {
                    'label': text,
                    'count': 0,
                }
            stats[key]['count'] += int(weight)

        preferred_text = str(preferred_category or '').strip()
        if preferred_text:
            # bevorzugte Kategorie ist immer verfügbar, aber zählt nicht künstlich hoch
            _add(preferred_text, weight=0)

        # Aus aktuell geladenen Aufgaben (reale Nutzung)
        for task in (self.loaded_tasks or []):
            _add(task.get('category'), weight=1)

        # Aus aktueller Session (reale Nutzung)
        if self.import_session:
            for task in getattr(self.import_session, 'tasks', []) or []:
                _add(getattr(task, 'category', ''), weight=1)

        # Aus konfigurierten Kategorie-Defaults (nur Verfügbarkeit, kein Nutzungsboost)
        if isinstance(category_defaults, dict):
            for key in category_defaults.keys():
                _add(key, weight=0)

        if not stats:
            return [preferred_text] if preferred_text else []

        preferred_key = preferred_text.lower() if preferred_text else None
        ranked_items = [
            item for key, item in stats.items()
            if key != preferred_key
        ]

        ranked_items.sort(key=lambda item: (-int(item.get('count', 0)), str(item.get('label', '')).lower()))

        suggestions = []
        if preferred_key and preferred_key in stats:
            suggestions.append(str(stats[preferred_key].get('label', preferred_text)))

        suggestions.extend(str(item.get('label', '')).strip() for item in ranked_items if str(item.get('label', '')).strip())
        return suggestions

    def _ask_import_metadata_dialog(self, initial_values, allowed_values):
        """Zeigt einen strukturierten Metadaten-Dialog mit Live-Validierung."""
        result: dict[str, Any] = {'value': None}
        hints = self._load_task_authoring_hints()
        field_hints = hints.get('field_hints', {}) or {}
        category_defaults = hints.get('category_defaults', {}) or {}

        def _hint_for(key, fallback=''):
            text = str(field_hints.get(key, '') or '').strip()
            return text or fallback

        def _category_key(value):
            return str(value or '').strip().lower()

        def _keyword_csv(value):
            if isinstance(value, str):
                return value.strip()
            if isinstance(value, (list, tuple, set)):
                items = [str(item).strip() for item in value if str(item).strip()]
                return ', '.join(items)
            return ''

        dialog = tk.Toplevel(self.root)
        dialog.title("Metadaten für neue Aufgabe")
        dialog.transient(self.root)
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding="12")
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(
            frame,
            text="Bitte Metadaten vollständig pflegen. Pflicht: Kategorie und Schwierigkeitsgrad.",
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 8))

        category_var = tk.StringVar(value=str(initial_values.get('category', '') or ''))
        difficulty_var = tk.StringVar(value=str(initial_values.get('difficulty', '') or ''))
        keywords_var = tk.StringVar(value=str(initial_values.get('keywords', '') or ''))
        title_var = tk.StringVar(value=str(initial_values.get('title', '') or ''))
        category_suggestion_var = tk.StringVar(value="")

        category_values = self._category_suggestions(
            preferred_category=category_var.get(),
            category_defaults=category_defaults,
        )

        ttk.Label(frame, text="Kategorie *").grid(row=1, column=0, sticky=tk.W, pady=(0, 6))
        category_combo = ttk.Combobox(
            frame,
            textvariable=category_var,
            values=category_values,
            state="normal",
            width=46,
        )
        category_combo.grid(row=1, column=1, columnspan=2, sticky="we", pady=(0, 6))
        ttk.Label(
            frame,
            text=_hint_for('category', 'Beispiel: Netzwerktechnik, IT-Sicherheit, Office-Praxis'),
        ).grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=(0, 8))
        ttk.Label(frame, textvariable=category_suggestion_var).grid(row=2, column=1, columnspan=2, sticky=tk.E, pady=(0, 8))

        ttk.Label(frame, text="Schwierigkeitsgrad *").grid(row=3, column=0, sticky=tk.W, pady=(0, 6))
        difficulty_combo = ttk.Combobox(
            frame,
            textvariable=difficulty_var,
            values=list(allowed_values),
            state="readonly",
            width=20,
        )
        difficulty_combo.grid(row=3, column=1, sticky=tk.W, pady=(0, 6))
        ttk.Label(
            frame,
            text=_hint_for('difficulty', 'Erlaubte Werte: leicht | mittel | schwer'),
        ).grid(row=4, column=1, columnspan=2, sticky=tk.W, pady=(0, 8))

        ttk.Label(frame, text="Schlagworte").grid(row=5, column=0, sticky=tk.W, pady=(0, 6))
        ttk.Entry(frame, textvariable=keywords_var, width=48).grid(row=5, column=1, columnspan=2, sticky="we", pady=(0, 6))
        ttk.Label(
            frame,
            text=_hint_for('keywords', 'Beispiel: VLAN, Routing, Fehlersuche'),
        ).grid(row=6, column=1, columnspan=2, sticky=tk.W, pady=(0, 8))

        ttk.Label(frame, text="Titel (optional)").grid(row=7, column=0, sticky=tk.W, pady=(0, 6))
        ttk.Entry(frame, textvariable=title_var, width=48).grid(row=7, column=1, columnspan=2, sticky="we", pady=(0, 6))
        ttk.Label(
            frame,
            text=_hint_for('title', "Kurz und konkret, z. B. 'Subnetzplanung Filiale Nord'"),
        ).grid(row=8, column=1, columnspan=2, sticky=tk.W, pady=(0, 8))

        status_var = tk.StringVar(value="")
        status_label = ttk.Label(frame, textvariable=status_var)
        status_label.grid(row=9, column=0, columnspan=3, sticky=tk.W, pady=(4, 10))

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=10, column=0, columnspan=3, sticky=tk.E)

        def _apply_category_defaults(force=False):
            cfg = category_defaults.get(_category_key(category_var.get()))
            if not isinstance(cfg, dict):
                category_suggestion_var.set("")
                return

            suggested_keywords = _keyword_csv(cfg.get('keywords'))
            suggested_difficulty = self._normalize_difficulty_value(cfg.get('difficulty'))
            title_prefix = str(cfg.get('title_prefix', '') or '').strip()

            suggestion_parts = []
            if suggested_keywords:
                suggestion_parts.append(f"KW: {suggested_keywords}")
            if suggested_difficulty:
                suggestion_parts.append(f"Diff: {suggested_difficulty}")
            if title_prefix:
                suggestion_parts.append(f"Titel-Präfix: {title_prefix}")

            category_suggestion_var.set(" | ".join(suggestion_parts))

            if suggested_keywords and (force or not str(keywords_var.get() or '').strip()):
                keywords_var.set(suggested_keywords)

            if suggested_difficulty:
                current_diff = self._normalize_difficulty_value(difficulty_var.get())
                if force or not current_diff:
                    difficulty_var.set(suggested_difficulty)

            if title_prefix and (force or not str(title_var.get() or '').strip()):
                title_var.set(title_prefix)

        def _validate():
            category_ok = bool(str(category_var.get() or '').strip())
            difficulty_ok = self._normalize_difficulty_value(difficulty_var.get()) is not None

            if category_ok and difficulty_ok:
                status_var.set("✓ Eingaben sind gültig.")
                ok_btn.configure(state=tk.NORMAL)
                return True

            reasons = []
            if not category_ok:
                reasons.append("Kategorie fehlt")
            if not difficulty_ok:
                reasons.append("Schwierigkeitsgrad ungültig")

            status_var.set("⚠ " + " | ".join(reasons))
            ok_btn.configure(state=tk.DISABLED)
            return False

        def _on_save():
            if not _validate():
                return

            normalized_difficulty = self._normalize_difficulty_value(difficulty_var.get())
            selected_category = str(category_var.get() or '').strip()
            result['value'] = {
                'category': selected_category,
                'difficulty': normalized_difficulty,
                'keywords': str(keywords_var.get() or '').strip(),
                'title': str(title_var.get() or '').strip(),
            }
            self._set_last_used_category(selected_category)
            dialog.destroy()

        def _on_cancel():
            result['value'] = None
            dialog.destroy()

        ttk.Button(
            btn_frame,
            text="Hilfe anzeigen",
            command=self.show_task_authoring_help,
        ).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(
            btn_frame,
            text="Vorschläge übernehmen",
            command=lambda: _apply_category_defaults(force=True),
        ).pack(side=tk.LEFT, padx=(0, 8))

        ok_btn = ttk.Button(btn_frame, text="Übernehmen", command=_on_save)
        ok_btn.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="Abbrechen", command=_on_cancel).pack(side=tk.LEFT)

        for var in (category_var, difficulty_var, keywords_var, title_var):
            var.trace_add('write', lambda *_args: _validate())

        category_var.trace_add('write', lambda *_args: _apply_category_defaults(force=False))

        dialog.protocol("WM_DELETE_WINDOW", _on_cancel)
        dialog.bind('<Escape>', lambda _e: _on_cancel())
        dialog.bind('<Return>', lambda _e: _on_save())

        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(2, weight=1)

        if not str(difficulty_var.get() or '').strip() and allowed_values:
            difficulty_var.set(str(allowed_values[0]))

        _apply_category_defaults(force=False)
        _validate()
        category_combo.focus_set()

        dialog.grab_set()
        self.root.wait_window(dialog)
        return result['value']

    def _ask_import_metadata(self, default_category=None, default_difficulty=None, default_keywords=None, default_title=None):
        """Fragt Metadaten für Aufgabenimport ab und gibt diese normalisiert zurück."""
        if not self._authoring_help_prompted:
            show_help = messagebox.askyesno(
                "Eingabehilfe anzeigen",
                "Möchten Sie vor der Metadaten-Eingabe eine kurze Hilfe mit JSON-Hinweisen und Formulierungsbeispielen sehen?",
            )
            self._authoring_help_prompted = True
            if show_help:
                self.show_task_authoring_help()

        rules_default_category = str(self._rule_value('default_import_metadata.category', '') or '')
        rules_default_difficulty = str(self._rule_value('default_import_metadata.difficulty', 'mittel') or 'mittel')
        rules_default_keywords = str(self._rule_value('default_import_metadata.keywords', '') or '')
        rules_default_title = str(self._rule_value('default_import_metadata.title', '') or '')
        allowed_values = self._difficulty_allowed_values()
        last_used_category = self._get_last_used_category()

        category_default = (
            default_category
            if default_category is not None
            else (last_used_category or self.lek_theme or rules_default_category or "Allgemein")
        )
        difficulty_default_raw = default_difficulty if default_difficulty is not None else rules_default_difficulty
        difficulty_default = self._normalize_difficulty_value(difficulty_default_raw) or (allowed_values[0] if allowed_values else 'mittel')
        keywords_default = default_keywords if default_keywords is not None else rules_default_keywords
        title_default = default_title if default_title is not None else rules_default_title

        initial_values = {
            'category': category_default,
            'difficulty': difficulty_default,
            'keywords': keywords_default,
            'title': title_default,
        }
        return self._ask_import_metadata_dialog(initial_values, allowed_values)

    def _rule_value(self, key, default=None):
        """Liest Import-Regelwerte über den WordProcessor (inkl. Punktnotation)."""
        getter = getattr(self.word_processor, 'get_import_rule', None)
        if callable(getter):
            return getter(key, default)
        return default

    def _preview_and_confirm_import(self, source_file, target_collection, category, difficulty, keywords, title='', allow_cancel=False):
        """Zeigt Preview/Details und liefert Import-Entscheidung zurück.

        Returns:
            bool|None: True=importieren, False=überspringen, None=abbrechen (nur wenn allow_cancel=True)
        """
        try:
            preview = self.word_processor.preview_task_append(
                source_doc_path=source_file,
                target_collection_path=target_collection,
                category=category,
                difficulty=difficulty,
                keywords=keywords,
                title=title,
            )
        except Exception as e:
            messagebox.showerror("Fehler", f"Vorschau fehlgeschlagen ({os.path.basename(source_file)}): {str(e)}")
            return False

        if preview.get('difficulty_inconsistent'):
            messagebox.showwarning(
                "Übernahme blockiert",
                "Der eingegebene Schwierigkeitsgrad ist inkonsistent (mehrere Werte).\n\n"
                "Bitte genau einen Wert verwenden: leicht, mittel oder schwer.",
            )
            return None if allow_cancel else False

        duplicate_check = preview.get('duplicate_check') or {}
        if duplicate_check.get('is_duplicate'):
            similarity_pct = int((duplicate_check.get('similarity') or 0) * 100)
            duplicate_text = (
                "Duplikatverdacht erkannt.\n\n"
                f"Quelle: {os.path.basename(source_file)}\n"
                f"Ähnlichkeit: {similarity_pct}%\n"
                f"Treffer-ID: {duplicate_check.get('matched_id') or '-'}\n"
                f"Treffer-Titel: {duplicate_check.get('matched_title') or '-'}\n"
                f"Treffer-Kategorie: {duplicate_check.get('matched_category') or '-'}\n\n"
            )

            if allow_cancel:
                dup_answer = messagebox.askyesnocancel(
                    "Duplikatverdacht",
                    duplicate_text
                    + "Ja = trotzdem fortfahren\n"
                    + "Nein = diese Datei überspringen\n"
                    + "Abbrechen = Serie stoppen",
                )
                if dup_answer is None:
                    return None
                if dup_answer is False:
                    return False
            else:
                dup_answer = messagebox.askyesno(
                    "Duplikatverdacht",
                    duplicate_text + "Trotzdem fortfahren?",
                )
                if not dup_answer:
                    return False

        lines = preview.get('source_preview_lines', [])
        source_preview_text = "\n".join(f"- {line}" for line in lines) if lines else "- (kein Textvorschau verfügbar)"
        if preview.get('source_paragraph_count', 0) > len(lines):
            source_preview_text += f"\n- ... (+{preview['source_paragraph_count'] - len(lines)} weitere Absätze)"

        details_question = messagebox.askyesno(
            "Details anzeigen",
            f"Möchten Sie vor der Übernahme die Detailansicht sehen?\n\nQuelle: {os.path.basename(source_file)}",
        )
        if details_question:
            self._show_task_import_details(preview)

        title = "Übernahme bestätigen"
        text = (
            f"Quelle: {os.path.basename(source_file)}\n"
            "Neue Aufgabe wird in die Aufgabensammlung übernommen.\n\n"
            f"Ziel-ID: {preview.get('next_id', '-')}\n"
            f"Titel: {preview.get('title', '-') }\n"
            f"Kategorie: {preview.get('category', '-') }\n"
            f"Schwierigkeit: {preview.get('difficulty_normalized', '-') }\n"
            f"Schlagworte: {preview.get('keywords', '-') or '-'}\n"
            f"Quellstruktur: {preview.get('source_paragraph_count', 0)} Absätze, "
            f"{preview.get('source_table_count', 0)} Tabellen\n\n"
            "Quell-Vorschau:\n"
            f"{source_preview_text}\n\n"
        )

        if allow_cancel:
            answer = messagebox.askyesnocancel(
                title,
                text + "Ja = übernehmen | Nein = überspringen | Abbrechen = Serie stoppen",
            )
            return answer

        return messagebox.askyesno(title, text + "Fortfahren?")

    def _normalize_difficulty_value(self, value):
        """Normalisiert Benutzereingaben auf leicht/mittel/schwer oder liefert None."""
        val = str(value or '').strip().lower()
        if not val:
            return None

        allowed = set(self._difficulty_allowed_values())
        if val in allowed:
            return val

        aliases: Any = self._rule_value('difficulty_rules.aliases', {}) or {}
        if not isinstance(aliases, dict):
            aliases = {}
        merged_aliases = {}
        for key, mapped in aliases.items():
            k = str(key).strip().lower()
            v = str(mapped).strip().lower()
            if k and v:
                merged_aliases[k] = v

        mapped_val = merged_aliases.get(val)
        if mapped_val in allowed:
            return mapped_val
        return None

    def _difficulty_allowed_values(self):
        """Liefert erlaubte Difficulty-Werte als normalisierte Kleinbuchstaben."""
        values: Any = self._rule_value('difficulty_rules.allowed_values', ['leicht', 'mittel', 'schwer']) or []
        if not isinstance(values, (list, tuple, set)):
            values = []
        normalized = [str(v).strip().lower() for v in values if str(v).strip()]
        return normalized or ['leicht', 'mittel', 'schwer']

    def _show_task_import_details(self, preview):
        """Zeigt eine detaillierte Vorschau der Quellblöcke vor der Aufgabenübernahme."""
        blocks = preview.get('source_preview_blocks', []) or []
        if not blocks:
            messagebox.showinfo("Detailansicht", "Keine Detaildaten zur Quelle verfügbar.")
            return

        lines = []
        for idx, block in enumerate(blocks, 1):
            btype = block.get('type')
            if btype == 'paragraph':
                text = str(block.get('text', '')).strip().replace('\n', ' ')
                if len(text) > 120:
                    text = text[:117] + '...'
                lines.append(f"{idx}. Absatz: {text}")
            elif btype == 'table':
                rows = block.get('rows', 0)
                cols = block.get('cols', 0)
                first_cell = str(block.get('first_cell', '')).strip()
                preview_cell = f" | erste Zelle: {first_cell}" if first_cell else ""
                lines.append(f"{idx}. Tabelle: {rows}x{cols}{preview_cell}")
            else:
                lines.append(f"{idx}. Unbekannter Blocktyp")

        details_text = "\n".join(lines)
        messagebox.showinfo(
            "Detailansicht Quellinhalt",
            "Erste Inhaltsblöcke der Quelle:\n\n" + details_text,
        )
    
    def _extract_lek_theme(self, filename):
        """
        Extrahiert das LEK-Thema aus dem Dateinamen
        
        Args:
            filename (str): Name der Datei (z.B. "Aufgaben_Mathematik.docx")
            
        Returns:
            str: Extrahiertes LEK-Thema (z.B. "Mathematik")
        """
        import re
        
        # Entferne die Dateierweiterung
        name_without_ext = re.sub(r'\.[^.]+$', '', filename)
        
        # Suche nach dem Muster "Aufgaben_XXX" 
        match = re.match(r'Aufgaben[_\-\s]*(.+)', name_without_ext, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        
        # Fallback: Verwende den gesamten Dateinamen ohne Extension
        return name_without_ext
    
    def filter_tasks(self, show_message=True):
        """Filtert Aufgaben basierend auf den angegebenen Kriterien"""
        if not self.loaded_tasks:
            messagebox.showwarning("Warnung", "Bitte laden Sie zuerst eine Aufgabensammlung.")
            return

        try:
            max_count = int(self.max_tasks_var.get())
        except Exception:
            max_count = 20
        max_count = max(1, min(20, max_count))
        self.max_tasks_var.set(str(max_count))
        
        criteria = {
            'keywords': [kw.strip() for kw in self.keywords_var.get().split(',') if kw.strip()],
            'difficulty': self.difficulty_var.get() if self.difficulty_var.get() != "Alle" else None,
            'max_count': max_count
        }
        
        filtered_tasks = self.task_selector.filter_tasks(self.loaded_tasks, criteria)
        blocker_mode = bool(self.only_blockers_var.get()) and self._is_advanced_mode()
        if blocker_mode:
            filtered_tasks = [task for task in filtered_tasks if self._task_has_blocking_warning(task)]

        self.current_displayed_tasks = filtered_tasks  # Speichere die gefilterten Aufgaben
        self.populate_task_tree(filtered_tasks)
        self._refresh_wizard_status()

        if show_message:
            mode_hint = " (nur Blocker)" if blocker_mode else ""
            messagebox.showinfo("Filter angewendet", f"{len(filtered_tasks)} Aufgaben entsprechen den Kriterien{mode_hint}.")
    
    def populate_task_tree(self, tasks):
        """Füllt die Treeview mit den gefundenen Aufgaben"""
        # Alle existierenden Einträge löschen
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        
        # Neue Einträge hinzufügen
        for i, task in enumerate(tasks, 1):
            # Interne, stabile ID für Auswahl/Freigabe
            internal_number = task.get('number', i)
            # Sichtbare Nummer inkl. Haupt-/Nebennummer (z. B. 1, 1.1, 1.2)
            display_number = self._task_number_label(task)
            
            # Keywords sicher formatieren
            keywords = task.get('keywords', [])
            keywords_text = ', '.join(keywords) if keywords else ''
            warnings = task.get('warnings', [])
            hints = task.get('hints', [])
            warnings_text = self._format_warning_preview(warnings, hints=hints, max_items=2)

            self.task_tree.insert("", "end", iid=str(internal_number), values=(
                display_number,
                task.get('category', 'Ohne Kategorie'),
                task.get('title', 'Ohne Titel'),
                task.get('difficulty', 'Unbekannt'),
                warnings_text,
                keywords_text
            ))

        self._update_tree_heading_sort_indicator()

    def _update_tree_heading_sort_indicator(self):
        """Aktualisiert die Header-Texte inkl. Sortierpfeil für die aktive Spalte."""
        columns = ("Nr", "Kategorie", "Titel", "Schwierigkeit", "Hinweise", "Suchbegriffe")
        for column in columns:
            arrow = ""
            if self._sort_column == column:
                arrow = " ▲" if not self._sort_reverse else " ▼"
            self.task_tree.heading(column, text=f"{column}{arrow}")

    def _number_sort_key(self, value):
        """Erstellt einen stabilen Sortierschlüssel für Nummern wie 1, 1.1, 2.10."""
        text = str(value or '').strip()
        if not text:
            return (float('inf'), float('inf'), text)

        parts = text.split('.')
        nums = []
        for part in parts[:2]:
            try:
                nums.append(int(part))
            except Exception:
                nums.append(float('inf'))

        while len(nums) < 2:
            nums.append(float('inf'))

        return (nums[0], nums[1], text.lower())

    def _task_sort_key(self, task, column):
        """Liefert den Sortierschlüssel je angeklickter Spalte."""
        if column == "Nr":
            return self._number_sort_key(self._task_number_label(task))
        if column == "Kategorie":
            return str(task.get('category') or '').strip().lower()
        if column == "Titel":
            return str(task.get('title') or '').strip().lower()
        if column == "Schwierigkeit":
            difficulty_order = {'leicht': 0, 'mittel': 1, 'schwer': 2}
            val = str(task.get('difficulty') or '').strip().lower()
            return (difficulty_order.get(val, 99), val)
        if column == "Hinweise":
            warnings = list(task.get('warnings', []) or [])
            text = self._format_warning_preview(warnings, max_items=5)
            return (0 if warnings else 1, text.lower())
        if column == "Suchbegriffe":
            keywords = ', '.join(task.get('keywords', []) or [])
            return keywords.lower()
        return str(task.get('title') or '').strip().lower()

    def _sort_task_tree_by_column(self, column):
        """Sortiert die aktuell angezeigten Aufgaben nach der angegebenen Spalte."""
        if not self.current_displayed_tasks:
            return

        if self._sort_column == column:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_column = column
            self._sort_reverse = False

        self.current_displayed_tasks = sorted(
            self.current_displayed_tasks,
            key=lambda task: self._task_sort_key(task, column),
            reverse=self._sort_reverse,
        )
        self.populate_task_tree(self.current_displayed_tasks)

    def _sync_import_session_task(self, edited_task):
        """Synchronisiert geänderte Task-Metadaten in die ImportSession."""
        if not self.import_session or not isinstance(edited_task, dict):
            return

        task_id = edited_task.get('number')
        for session_task in self.import_session.tasks:
            if session_task.id != task_id:
                continue

            session_task.raw_task = edited_task
            session_task.category = str(edited_task.get('category') or '').strip()
            session_task.title = str(edited_task.get('title') or '').strip()
            session_task.difficulty = str(edited_task.get('difficulty') or '').strip()
            session_task.keywords = list(edited_task.get('keywords', []) or [])
            session_task.warnings = list(edited_task.get('warnings', []) or [])
            session_task.intro = list(edited_task.get('intro', []) or [])
            break

    def _on_task_double_click(self, event):
        """Öffnet die Bearbeitung für die doppelt geklickte Aufgabe."""
        row_id = self.task_tree.identify_row(event.y)
        if not row_id:
            return

        try:
            internal_id = int(row_id)
        except Exception:
            return

        task = self._find_task_by_internal_number(internal_id)
        if not task:
            return

        self._open_task_edit_dialog(task)

    def _open_task_edit_dialog(self, task):
        """Bearbeitet Metadaten und Inhalt einer Aufgabe in einem Dialog."""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Aufgabe bearbeiten – Nr {self._task_number_label(task)}")
        dialog.transient(self.root)
        dialog.geometry("920x700")

        frame = ttk.Frame(dialog, padding="12")
        frame.pack(fill=tk.BOTH, expand=True)

        title_var = tk.StringVar(value=str(task.get('title') or ''))
        category_var = tk.StringVar(value=str(task.get('category') or ''))
        difficulty_var = tk.StringVar(value=str(task.get('difficulty') or ''))
        keywords_var = tk.StringVar(value=', '.join(task.get('keywords', []) or []))

        ttk.Label(frame, text="Titel *").grid(row=0, column=0, sticky=tk.W, pady=(0, 6))
        ttk.Entry(frame, textvariable=title_var, width=60).grid(row=0, column=1, sticky="we", pady=(0, 6))

        ttk.Label(frame, text="Kategorie *").grid(row=1, column=0, sticky=tk.W, pady=(0, 6))
        ttk.Entry(frame, textvariable=category_var, width=60).grid(row=1, column=1, sticky="we", pady=(0, 6))

        ttk.Label(frame, text="Schwierigkeit *").grid(row=2, column=0, sticky=tk.W, pady=(0, 6))
        difficulty_combo = ttk.Combobox(
            frame,
            textvariable=difficulty_var,
            values=self._difficulty_allowed_values(),
            state="readonly",
            width=20,
        )
        difficulty_combo.grid(row=2, column=1, sticky=tk.W, pady=(0, 6))

        ttk.Label(frame, text="Schlagworte (kommagetrennt)").grid(row=3, column=0, sticky=tk.W, pady=(0, 6))
        ttk.Entry(frame, textvariable=keywords_var, width=60).grid(row=3, column=1, sticky="we", pady=(0, 6))

        ttk.Label(frame, text="Inhalt (zeilenweise)").grid(row=4, column=0, sticky=tk.W, pady=(8, 6))
        content_text = scrolledtext.ScrolledText(frame, wrap=tk.WORD, height=18, font=("Segoe UI", 10))
        content_text.grid(row=4, column=1, sticky="nsew", pady=(8, 6))
        existing_content = task.get('content', []) or []
        content_text.insert(tk.END, "\n".join(str(line) for line in existing_content))

        status_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=status_var).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(2, 8))

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=6, column=0, columnspan=2, sticky=tk.E)

        def _validate_inputs():
            title_ok = bool(str(title_var.get() or '').strip())
            category_ok = bool(str(category_var.get() or '').strip())
            difficulty_ok = self._normalize_difficulty_value(difficulty_var.get()) is not None

            if title_ok and category_ok and difficulty_ok:
                status_var.set("✓ Eingaben sind gültig.")
                save_btn.configure(state=tk.NORMAL)
                return True

            reasons = []
            if not title_ok:
                reasons.append("Titel fehlt")
            if not category_ok:
                reasons.append("Kategorie fehlt")
            if not difficulty_ok:
                reasons.append("Schwierigkeit ungültig")
            status_var.set("⚠ " + " | ".join(reasons))
            save_btn.configure(state=tk.DISABLED)
            return False

        def _on_save():
            if not _validate_inputs():
                return

            new_keywords = [k.strip() for k in str(keywords_var.get() or '').split(',') if k.strip()]
            raw_lines = [line.rstrip() for line in content_text.get('1.0', tk.END).splitlines()]
            # Leere Absätze bewusst erhalten (nur überflüssige Leerzeilen am Ende entfernen)
            while raw_lines and raw_lines[-1] == '':
                raw_lines.pop()
            new_content = raw_lines

            task['title'] = str(title_var.get() or '').strip()
            task['category'] = str(category_var.get() or '').strip()
            task['difficulty'] = self._normalize_difficulty_value(difficulty_var.get()) or str(difficulty_var.get() or '').strip()
            task['keywords'] = new_keywords
            task['content'] = new_content

            # Diagnostik nach Bearbeitung neu berechnen (wenn verfügbar)
            try:
                diagnostic = self.word_processor._build_task_diagnostic(task)
                task['intro'] = diagnostic.get('intro', [])
                task['warnings'] = diagnostic.get('warnings', [])
                task['hints'] = diagnostic.get('hints', [])
            except Exception:
                pass

            self._sync_import_session_task(task)
            self.populate_task_tree(self.current_displayed_tasks)
            self._refresh_wizard_status()
            dialog.destroy()
            messagebox.showinfo("Aufgabe aktualisiert", "Die Aufgabe wurde erfolgreich aktualisiert.")

        def _on_cancel():
            dialog.destroy()

        save_btn = ttk.Button(btn_frame, text="Speichern", command=_on_save)
        save_btn.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="Abbrechen", command=_on_cancel).pack(side=tk.LEFT)

        for var in (title_var, category_var, difficulty_var, keywords_var):
            var.trace_add('write', lambda *_args: _validate_inputs())

        _validate_inputs()

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(4, weight=1)

        dialog.grab_set()
        dialog.focus_force()

    def _task_number_label(self, task):
        """Liefert die anzuzeigende Aufgaben-Nummer (Haupt-/Nebennummer), falls vorhanden."""
        return str(
            task.get('number_display')
            or task.get('number_label')
            or task.get('number')
            or '-'
        )

    def _selected_task_ids_from_tree(self):
        """Liefert die aktuell markierten Aufgaben-IDs aus der Treeview."""
        selected_ids = []
        for item in self.task_tree.selection():
            try:
                selected_ids.append(int(item))
            except (TypeError, ValueError):
                continue
        return selected_ids

    def _task_group_key(self, task):
        """Ermittelt den Gruppenschlüssel für Haupt-/Unteraufgaben (z. B. 1.0/1.1 -> 1)."""
        label = self._task_number_label(task)
        text = str(label or '').strip()
        if text:
            match = re.match(r'^(\d+)(?:\s*[\.,]\s*\d+)?', text)
            if match:
                return str(match.group(1)).strip()

            leading_number = re.match(r'^(\d+)', text)
            if leading_number:
                return str(leading_number.group(1)).strip()

        internal_number = task.get('number')
        if internal_number is None:
            return ''
        return str(internal_number).strip()

    def _expand_related_task_ids(self, task_ids):
        """Erweitert eine Auswahl auf alle zugehörigen Haupt-/Unteraufgaben."""
        normalized_ids = []
        for task_id in task_ids or []:
            try:
                normalized_ids.append(int(task_id))
            except (TypeError, ValueError):
                continue

        if not normalized_ids:
            return []

        if self.import_session:
            related_ids = self.import_session.get_related_task_ids(normalized_ids)
            return [int(task_id) for task_id in related_ids]

        all_tasks = list(getattr(self, 'loaded_tasks', []) or [])
        if not all_tasks:
            return normalized_ids

        selected_group_keys = set()
        for task in all_tasks:
            task_number = task.get('number')
            if task_number in normalized_ids:
                group_key = self._task_group_key(task)
                if group_key:
                    selected_group_keys.add(group_key)

        if not selected_group_keys:
            return normalized_ids

        return [
            int(task.get('number'))
            for task in all_tasks
            if task.get('number') is not None and self._task_group_key(task) in selected_group_keys
        ]

    def _find_task_by_internal_number(self, internal_number):
        """Sucht eine Aufgabe über die interne Nummer in der aktuell angezeigten Liste."""
        current_displayed_tasks = getattr(self, 'current_displayed_tasks', self.loaded_tasks)
        for task in current_displayed_tasks:
            if task.get('number', 0) == internal_number:
                return task
        return None

    def preview_selected_task(self):
        """Zeigt eine inhaltliche Vorschau der aktuell markierten Aufgabe als Fließtext an."""
        selected_ids = self._selected_task_ids_from_tree()
        if not selected_ids:
            messagebox.showwarning("Warnung", "Bitte mindestens eine Aufgabe auswählen.")
            return

        if len(selected_ids) > 1:
            messagebox.showinfo(
                "Hinweis",
                "Es wurden mehrere Aufgaben markiert. Für die Vorschau wird die erste Auswahl verwendet.",
            )

        task_id = selected_ids[0]
        task = self._find_task_by_internal_number(task_id)
        if not task:
            messagebox.showwarning("Hinweis", "Ausgewählte Aufgabe konnte nicht gefunden werden.")
            return

        title = str(task.get('title') or 'Ohne Titel').strip()
        category = str(task.get('category') or 'Ohne Kategorie').strip()
        difficulty = str(task.get('difficulty') or 'Unbekannt').strip()
        number_label = self._task_number_label(task)

        content_lines = [str(line).strip() for line in (task.get('content') or []) if str(line).strip()]
        keywords = task.get('keywords') or []
        warnings = task.get('warnings') or []

        text_lines = [
            f"Nr: {number_label}",
            f"Titel: {title}",
            f"Kategorie: {category}",
            f"Schwierigkeit: {difficulty}",
            "",
            "Inhalt:",
        ]

        if content_lines:
            for idx, line in enumerate(content_lines, 1):
                text_lines.append(f"{idx}. {line}")
        else:
            text_lines.append("(kein Inhalt verfügbar)")

        if keywords:
            text_lines.extend(["", "Schlagworte:", ", ".join(str(k) for k in keywords)])

        if warnings:
            text_lines.extend(["", "Warnungen:"])
            text_lines.extend(f"- {w}" for w in warnings)

        preview_text = "\n".join(text_lines)

        preview_window = tk.Toplevel(self.root)
        preview_window.title(f"Aufgaben-Vorschau #{number_label}")
        preview_window.geometry("900x620")

        text_widget = scrolledtext.ScrolledText(preview_window, wrap=tk.WORD, font=("Segoe UI", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        text_widget.insert(tk.END, preview_text)
        text_widget.configure(state=tk.DISABLED)

    def preview_lek_output(self):
        """Zeigt eine Gesamtvorschau der freigegebenen LEK-Ausgabe (Schritt 3)."""
        if not self.import_session:
            messagebox.showwarning("Warnung", "Keine Import-Session vorhanden. Bitte zuerst Datei laden.")
            return

        approved_tasks = self.import_session.get_approved_raw_tasks()
        if not approved_tasks:
            messagebox.showwarning(
                "Hinweis",
                "Keine freigegebenen Aufgaben vorhanden.\n"
                "Bitte zuerst Aufgaben auswählen und über 'Auswahl freigeben' bestätigen.",
            )
            return

        lines = [
            "LEK-Gesamtvorschau (Schritt 3)",
            "=" * 36,
        ]

        if str(self.lek_theme or '').strip():
            lines.extend([
                f"Thema: {self.lek_theme}",
                "",
            ])

        lines.extend([
            f"Freigegebene Aufgaben: {len(approved_tasks)}",
            "",
        ])

        preview_builder = getattr(self.word_processor, 'build_task_flow_preview_lines', None)
        delta_checker = getattr(self.word_processor, 'analyze_task_flow_preview_delta', None)

        for idx, task in enumerate(approved_tasks, 1):
            number_label = self._task_number_label(task)
            task_title = str(task.get('title') or 'Ohne Titel').strip()
            lines.append(f"{idx}) Nr {number_label} – {task_title}")

            section_lines = []
            if callable(preview_builder):
                try:
                    section_lines = preview_builder(task)
                except Exception:
                    section_lines = []

            if not isinstance(section_lines, list):
                section_lines = []

            if not section_lines:
                section_lines = [
                    str(line).strip()
                    for line in (task.get('content') or [])
                    if str(line).strip()
                ]
                if not section_lines:
                    section_lines = ['(kein Inhalt verfügbar)']

            missing_optionals = []
            if callable(delta_checker):
                try:
                    delta = delta_checker(task) or {}
                    if isinstance(delta, dict):
                        missing_optionals = list(delta.get('missing_optional_sections') or [])
                except Exception:
                    missing_optionals = []

            lines.append("")
            lines.extend(section_lines)

            if missing_optionals:
                lines.append("")
                lines.append(
                    "Delta-Check (optional nicht vorhanden): "
                    + ", ".join(str(item) for item in missing_optionals)
                )

            lines.append("")
            lines.append("-" * 72)
            lines.append("")

        preview_text = "\n".join(lines).rstrip() + "\n"

        preview_window = tk.Toplevel(self.root)
        preview_window.title("LEK-Gesamtvorschau")
        preview_window.geometry("980x720")

        text_widget = scrolledtext.ScrolledText(preview_window, wrap=tk.WORD, font=("Segoe UI", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        text_widget.insert(tk.END, preview_text)
        text_widget.configure(state=tk.DISABLED)

    def approve_selected_tasks(self):
        """Markiert aktuell ausgewählte Aufgaben als für den Export freigegeben."""
        if not self.import_session:
            messagebox.showwarning("Warnung", "Keine Import-Session vorhanden. Bitte zuerst Datei laden.")
            return

        selected_ids = self._selected_task_ids_from_tree()
        if not selected_ids:
            messagebox.showwarning("Warnung", "Bitte mindestens eine Aufgabe auswählen.")
            return

        effective_ids = self.import_session.set_task_approvals(selected_ids, True)

        stats = self.import_session.get_stats()
        self._set_step(3, show_message=False)
        self._refresh_wizard_status()

        effective_count = len(effective_ids)
        expanded_hint = ""
        if effective_count > len(selected_ids):
            expanded_hint = (
                f"\nDurch Haupt-/Unteraufgaben-Gruppierung wurden insgesamt {effective_count} Aufgabe(n) freigegeben."
            )

        messagebox.showinfo(
            "Freigabe aktualisiert",
            f"{len(selected_ids)} Aufgabe(n) freigegeben.\n"
            f"Aktuell freigegeben: {stats['approved']} von {stats['total']}."
            f"{expanded_hint}",
        )

    def clear_task_approvals(self):
        """Entfernt alle aktuellen Aufgaben-Freigaben."""
        if not self.import_session:
            return

        self.import_session.clear_approvals()
        if self.current_step > 3:
            self._set_step(3, show_message=False)
        self._refresh_wizard_status()
        messagebox.showinfo("Freigaben gelöscht", "Alle Freigaben wurden entfernt.")
    
    def select_all_tasks(self):
        """Wählt alle Aufgaben in der Treeview aus"""
        for item in self.task_tree.get_children():
            self.task_tree.selection_add(item)

    def select_blocker_tasks(self):
        """Markiert nur Aufgaben mit blockierenden Warnungen in der aktuellen Ansicht."""
        if not self.current_displayed_tasks:
            messagebox.showwarning("Hinweis", "Keine Aufgaben in der aktuellen Ansicht.")
            return

        self.task_tree.selection_remove(self.task_tree.selection())

        selected_count = 0
        for task in self.current_displayed_tasks:
            if not self._task_has_blocking_warning(task):
                continue

            task_id = str(task.get('number', ''))
            if task_id and self.task_tree.exists(task_id):
                self.task_tree.selection_add(task_id)
                selected_count += 1

        if selected_count == 0:
            messagebox.showinfo("Blocker-Auswahl", "Keine blockerbehafteten Aufgaben in der aktuellen Ansicht gefunden.")
            return

        messagebox.showinfo("Blocker-Auswahl", f"{selected_count} blockerbehaftete Aufgabe(n) wurden markiert.")

    def bulk_edit_selected_tasks(self):
        """Bearbeitet Kategorie/Schwierigkeit/Schlagworte für mehrere ausgewählte Aufgaben gleichzeitig."""
        selected_ids = self._selected_task_ids_from_tree()
        if not selected_ids:
            messagebox.showwarning("Hinweis", "Bitte zuerst mindestens eine Aufgabe auswählen.")
            return

        selected_tasks = []
        for task_id in selected_ids:
            task = self._find_task_by_internal_number(task_id)
            if task:
                selected_tasks.append(task)

        if not selected_tasks:
            messagebox.showwarning("Hinweis", "Die ausgewählten Aufgaben konnten nicht aufgelöst werden.")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Mehrfachbearbeitung ({len(selected_tasks)} Aufgaben)")
        dialog.transient(self.root)
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding="12")
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(
            frame,
            text="Sie können Kategorie, Schwierigkeit und/oder Schlagworte für alle ausgewählten Aufgaben setzen.",
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

        set_category_var = tk.BooleanVar(value=False)
        category_value_var = tk.StringVar(value="")

        set_difficulty_var = tk.BooleanVar(value=False)
        difficulty_value_var = tk.StringVar(value=self._difficulty_allowed_values()[0])

        set_keywords_var = tk.BooleanVar(value=False)
        keywords_value_var = tk.StringVar(value="")
        keywords_mode_var = tk.StringVar(value="ersetzen")

        ttk.Checkbutton(
            frame,
            text="Kategorie setzen",
            variable=set_category_var,
        ).grid(row=1, column=0, sticky=tk.W, pady=(0, 6))
        category_entry = ttk.Entry(frame, textvariable=category_value_var, width=40)
        category_entry.grid(row=1, column=1, columnspan=2, sticky="we", pady=(0, 6))

        ttk.Checkbutton(
            frame,
            text="Schwierigkeit setzen",
            variable=set_difficulty_var,
        ).grid(row=2, column=0, sticky=tk.W, pady=(0, 6))
        difficulty_combo = ttk.Combobox(
            frame,
            textvariable=difficulty_value_var,
            values=self._difficulty_allowed_values(),
            state="readonly",
            width=20,
        )
        difficulty_combo.grid(row=2, column=1, sticky=tk.W, pady=(0, 6))

        ttk.Checkbutton(
            frame,
            text="Schlagworte bearbeiten",
            variable=set_keywords_var,
        ).grid(row=3, column=0, sticky=tk.W, pady=(0, 6))
        keywords_entry = ttk.Entry(frame, textvariable=keywords_value_var, width=40)
        keywords_entry.grid(row=3, column=1, sticky="we", pady=(0, 6))
        keywords_mode_combo = ttk.Combobox(
            frame,
            textvariable=keywords_mode_var,
            values=["ersetzen", "anhängen"],
            state="readonly",
            width=12,
        )
        keywords_mode_combo.grid(row=3, column=2, sticky=tk.W, pady=(0, 6))

        status_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=status_var).grid(row=4, column=0, columnspan=3, sticky=tk.W, pady=(4, 10))

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=5, column=0, columnspan=3, sticky=tk.E)

        def _validate():
            wants_category = bool(set_category_var.get())
            wants_difficulty = bool(set_difficulty_var.get())
            wants_keywords = bool(set_keywords_var.get())

            if not wants_category and not wants_difficulty and not wants_keywords:
                status_var.set("⚠ Bitte mindestens eine Änderung aktivieren.")
                apply_btn.configure(state=tk.DISABLED)
                return False

            if wants_category and not str(category_value_var.get() or '').strip():
                status_var.set("⚠ Kategorie ist aktiviert, aber leer.")
                apply_btn.configure(state=tk.DISABLED)
                return False

            if wants_difficulty and self._normalize_difficulty_value(difficulty_value_var.get()) is None:
                status_var.set("⚠ Ungültiger Schwierigkeitsgrad.")
                apply_btn.configure(state=tk.DISABLED)
                return False

            if wants_keywords and not str(keywords_value_var.get() or '').strip():
                status_var.set("⚠ Schlagworte sind aktiviert, aber leer.")
                apply_btn.configure(state=tk.DISABLED)
                return False

            if wants_keywords and keywords_mode_var.get() not in ("ersetzen", "anhängen"):
                status_var.set("⚠ Ungültiger Schlagwort-Modus.")
                apply_btn.configure(state=tk.DISABLED)
                return False

            status_var.set("✓ Änderungen können angewendet werden.")
            apply_btn.configure(state=tk.NORMAL)
            return True

        def _apply_changes():
            if not _validate():
                return

            new_category = str(category_value_var.get() or '').strip()
            new_difficulty = self._normalize_difficulty_value(difficulty_value_var.get())
            raw_keywords = str(keywords_value_var.get() or '').strip()
            new_keywords = [part.strip() for part in raw_keywords.split(',') if part.strip()]
            keywords_mode = keywords_mode_var.get()

            changed = 0
            for task in selected_tasks:
                if set_category_var.get():
                    task['category'] = new_category
                if set_difficulty_var.get() and new_difficulty:
                    task['difficulty'] = new_difficulty
                if set_keywords_var.get():
                    current_keywords = task.get('keywords', [])
                    if not isinstance(current_keywords, list):
                        current_keywords = []

                    if keywords_mode == "ersetzen":
                        task['keywords'] = new_keywords
                    else:
                        existing_lookup = {str(item).strip().lower() for item in current_keywords if str(item).strip()}
                        merged_keywords = list(current_keywords)
                        for kw in new_keywords:
                            if kw.lower() not in existing_lookup:
                                merged_keywords.append(kw)
                                existing_lookup.add(kw.lower())
                        task['keywords'] = merged_keywords

                try:
                    diagnostic = self.word_processor._build_task_diagnostic(task)
                    task['intro'] = diagnostic.get('intro', [])
                    task['warnings'] = diagnostic.get('warnings', [])
                    task['hints'] = diagnostic.get('hints', [])
                except Exception:
                    pass

                self._sync_import_session_task(task)
                changed += 1

            if self._sort_column:
                self.current_displayed_tasks = sorted(
                    self.current_displayed_tasks,
                    key=lambda task: self._task_sort_key(task, self._sort_column),
                    reverse=self._sort_reverse,
                )

            self.populate_task_tree(self.current_displayed_tasks)
            self._refresh_wizard_status()
            dialog.destroy()
            messagebox.showinfo("Mehrfachbearbeitung", f"{changed} Aufgabe(n) wurden aktualisiert.")

        def _cancel():
            dialog.destroy()

        apply_btn = ttk.Button(btn_frame, text="Änderungen anwenden", command=_apply_changes)
        apply_btn.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="Abbrechen", command=_cancel).pack(side=tk.LEFT)

        for var in (
            set_category_var,
            category_value_var,
            set_difficulty_var,
            difficulty_value_var,
            set_keywords_var,
            keywords_value_var,
            keywords_mode_var,
        ):
            var.trace_add('write', lambda *_args: _validate())

        _validate()
        category_entry.focus_set()
        dialog.grab_set()
    
    def deselect_all_tasks(self):
        """Hebt die Auswahl aller Aufgaben auf"""
        self.task_tree.selection_remove(self.task_tree.selection())
    
    def export_selected(self):
        """Exportiert nur die markierten Aufgaben in eine neue Word-Datei"""
        if not self._set_step(4, show_message=True):
            return

        selected_items = self.task_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warnung", "Bitte wählen Sie mindestens eine Aufgabe aus.")
            return
        
        selected_ids = self._selected_task_ids_from_tree()
        selected_tasks = []
        effective_ids = self._expand_related_task_ids(selected_ids)

        # Bei vorhandener Session: Auswahl = Freigabe + Export aus bestätigten Tasks
        if self.import_session:
            self.import_session.set_task_approvals(selected_ids, True)
            selected_tasks = self.import_session.get_related_raw_tasks(selected_ids)
        else:
            # Fallback ohne Session
            all_tasks = list(getattr(self, 'loaded_tasks', []) or [])
            for task_id in effective_ids:
                for task in all_tasks:
                    if task.get('number', 0) == task_id:
                        selected_tasks.append(task)
                        break
        
        if not selected_tasks:
            messagebox.showwarning("Warnung", "Keine gültigen Aufgaben zur Auswahl gefunden.")
            return
        
        # Export durchführen
        self._perform_export(selected_tasks, f"{len(selected_tasks)} markierte Aufgabe(n)")
    
    def export_all(self):
        """Exportiert alle aktuell angezeigten Aufgaben in eine neue Word-Datei"""
        if not self._set_step(4, show_message=True):
            return

        if self.import_session:
            approved_tasks = self.import_session.get_approved_raw_tasks()
            if not approved_tasks:
                messagebox.showwarning(
                    "Warnung",
                    "Keine freigegebenen Aufgaben vorhanden.\n"
                    "Bitte Aufgaben auswählen und über 'Auswahl freigeben' bestätigen.",
                )
                return

            self._perform_export(approved_tasks, f"alle {len(approved_tasks)} freigegebene Aufgabe(n)")
            return

        current_displayed_tasks = getattr(self, 'current_displayed_tasks', self.loaded_tasks)
        if not current_displayed_tasks:
            messagebox.showwarning("Warnung", "Keine Aufgaben zum Exportieren vorhanden.")
            return

        # Fallback ohne Session
        self._perform_export(current_displayed_tasks, f"alle {len(current_displayed_tasks)} Aufgabe(n)")
    
    def _perform_export(self, tasks_to_export, description):
        """
        Führt den eigentlichen Export durch
        
        Args:
            tasks_to_export: Liste der zu exportierenden Aufgaben
            description: Beschreibung für die Erfolgsmeldung
        """
        # Verbindliche Vorabprüfung: Inkonsistenzen müssen vor Export bereinigt werden
        blocked_difficulty_tasks = []
        blocked_category_tasks = []
        blocked_required_tasks = []
        blocked_external_table_tasks = []
        blocked_duplicate_number_tasks = []
        block_difficulty = bool(self._rule_value('difficulty_rules.block_export_on_inconsistent', True))
        allowed_difficulty = {str(v).strip().lower() for v in self._difficulty_allowed_values()}
        block_category = bool(self._rule_value('category_rules.block_export_on_missing', True))
        block_required = bool(self._rule_value('template_rules.block_export_on_missing_required', True))
        block_external_table = bool(self._rule_value('external_table_rules.block_export_on_missing_reference', True))
        missing_values: Any = self._rule_value('category_rules.missing_values', ['', 'ohne kategorie']) or []
        if not isinstance(missing_values, (list, tuple, set)):
            missing_values = []
        missing_values_norm = {str(v).strip().lower() for v in missing_values}

        for task in tasks_to_export:
            warnings = task.get('warnings', []) or []
            if block_difficulty:
                difficulty_warning = any(
                    ('Inkonsistenter Schwierigkeitsgrad' in str(w)) or ('Schwierigkeit ist nicht eindeutig' in str(w))
                    for w in warnings
                )
                difficulty_value = str(task.get('difficulty', '') or '').strip().lower()
                difficulty_invalid = bool(difficulty_value) and difficulty_value not in allowed_difficulty
                if difficulty_warning or difficulty_invalid:
                    blocked_difficulty_tasks.append(
                        f"#{self._task_number_label(task)} {task.get('title', 'Ohne Titel')}"
                    )

            if block_category:
                category_value = str(task.get('category', '') or '').strip().lower()
                category_warning = any('Kategorie fehlt' in str(w) for w in warnings)
                if category_warning or category_value in missing_values_norm:
                    blocked_category_tasks.append(
                        f"#{self._task_number_label(task)} {task.get('title', 'Ohne Titel')}"
                    )

            if block_required:
                required_warning = any('Pflichtfeld fehlt:' in str(w) for w in warnings)
                if required_warning:
                    blocked_required_tasks.append(
                        f"#{self._task_number_label(task)} {task.get('title', 'Ohne Titel')}"
                    )

            if block_external_table:
                if bool(task.get('external_table_reference')) and bool(task.get('external_table_missing')):
                    blocked_external_table_tasks.append(
                        f"#{self._task_number_label(task)} {task.get('title', 'Ohne Titel')}"
                    )

            duplicate_warning = any('Doppelte Aufgabennummer erkannt' in str(w) for w in warnings)
            if duplicate_warning:
                blocked_duplicate_number_tasks.append(
                    f"#{self._task_number_label(task)} {task.get('title', 'Ohne Titel')}"
                )

        if blocked_difficulty_tasks or blocked_category_tasks or blocked_required_tasks or blocked_external_table_tasks or blocked_duplicate_number_tasks:
            details = []

            if blocked_difficulty_tasks:
                sample_diff = "\n".join(f"- {entry}" for entry in blocked_difficulty_tasks[:8])
                if len(blocked_difficulty_tasks) > 8:
                    sample_diff += f"\n- ... (+{len(blocked_difficulty_tasks) - 8} weitere)"
                details.append(
                    "Inkonsistenter Schwierigkeitsgrad:\n"
                    f"{sample_diff}"
                )

            if blocked_category_tasks:
                sample_cat = "\n".join(f"- {entry}" for entry in blocked_category_tasks[:8])
                if len(blocked_category_tasks) > 8:
                    sample_cat += f"\n- ... (+{len(blocked_category_tasks) - 8} weitere)"
                details.append(
                    "Fehlende Kategorie (Pflichtfeld):\n"
                    f"{sample_cat}"
                )

            if blocked_required_tasks:
                sample_required = "\n".join(f"- {entry}" for entry in blocked_required_tasks[:8])
                if len(blocked_required_tasks) > 8:
                    sample_required += f"\n- ... (+{len(blocked_required_tasks) - 8} weitere)"
                details.append(
                    "Fehlende Template-Pflichtfelder (z. B. Titel):\n"
                    f"{sample_required}"
                )

            if blocked_external_table_tasks:
                sample_external = "\n".join(f"- {entry}" for entry in blocked_external_table_tasks[:8])
                if len(blocked_external_table_tasks) > 8:
                    sample_external += f"\n- ... (+{len(blocked_external_table_tasks) - 8} weitere)"
                details.append(
                    "Fehlende externe Tabellenreferenz:\n"
                    f"{sample_external}"
                )

            if blocked_duplicate_number_tasks:
                sample_duplicates = "\n".join(f"- {entry}" for entry in blocked_duplicate_number_tasks[:8])
                if len(blocked_duplicate_number_tasks) > 8:
                    sample_duplicates += f"\n- ... (+{len(blocked_duplicate_number_tasks) - 8} weitere)"
                details.append(
                    "Doppelte Aufgabennummern:\n"
                    f"{sample_duplicates}"
                )

            messagebox.showwarning(
                "Export blockiert",
                "Der Export wurde gestoppt, weil Pflichtprüfungen fehlgeschlagen sind.\n\n"
                "Bitte bereinigen Sie die betroffenen Aufgaben direkt in der Quelldatei "
                "und laden Sie diese erneut.\n\n"
                + "\n\n".join(details),
            )
            return

        # Generiere Standarddateinamen
        from datetime import datetime
        now = datetime.now()
        year_month_day = now.strftime("%Y%m%d")
        time_stamp = now.strftime("%H%M")
        
        # LEK-Export-Verzeichnis erstellen  
        base_dir = _runtime_base_dir()
        leks_dir = base_dir / "data" / "LEKs"
        leks_dir.mkdir(parents=True, exist_ok=True)
        
        if self.lek_theme:
            default_filename = f"LEK_{self.lek_theme}_{year_month_day}_{time_stamp}.docx"
        else:
            default_filename = f"LEK_Aufgaben_{year_month_day}_{time_stamp}.docx"

        # Verbindliche Vorschau vor Export (Vorschau == Export)
        preview_titles = [t.get('title', 'Ohne Titel') for t in tasks_to_export[:5]]
        preview_text = "\n".join([f"- {title}" for title in preview_titles])
        if len(tasks_to_export) > 5:
            preview_text += f"\n- ... (+{len(tasks_to_export) - 5} weitere)"

        proceed = messagebox.askyesno(
            "Export bestätigen",
            f"Es werden {len(tasks_to_export)} Aufgabe(n) exportiert:\n\n{preview_text}\n\nJetzt speichern?",
        )
        if not proceed:
            return
        
        # Speicherpfad wählen
        output_file = filedialog.asksaveasfilename(
            title="Word-Datei speichern",
            defaultextension=".docx",
            initialfile=default_filename,
            initialdir=str(leks_dir),
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")]
        )
        
        if output_file:
            try:
                # LEK-Thema an WordProcessor weitergeben
                self.word_processor.create_document_from_tasks(tasks_to_export, output_file, self.lek_theme)
                messagebox.showinfo("Erfolg", f"{description} erfolgreich in '{output_file}' gespeichert.")
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler beim Speichern: {str(e)}")
    
    def export_selection(self):
        """Exportiert die ausgewählten Aufgaben in eine neue Word-Datei (Legacy-Methode)"""
        # Für Rückwärtskompatibilität - ruft export_all auf
        self.export_all()

def main():
    root = tk.Tk()
    app = LEKBastlerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()