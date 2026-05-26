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
import os
import sys
from pathlib import Path
from word_processor import WordProcessor
from task_selector import TaskSelector
from import_wizard import ImportSession


def _runtime_base_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]

class LEKBastlerGUI:
    def __init__(self, root):
        self.root = root
        self._apply_window_icon()
        self.root.title("LEK-Bastler - Aufgabenauswahl")
        self.root.geometry("800x600")
        self._apply_window_state()
        
        self.word_processor = WordProcessor()
        self.task_selector = TaskSelector()
        self.loaded_tasks = []
        self.current_displayed_tasks = []  # Speichert die aktuell angezeigten Aufgaben
        self.source_filename = ""  # Speichert den Namen der Quelldatei
        self.lek_theme = ""  # Speichert das extrahierte LEK-Thema
        self.import_session = None  # Wizard-Session (Sprint 1)
        self.current_step = 1
        self.step_labels = {
            1: "Quelle wählen",
            2: "Erkennung prüfen",
            3: "Aufgaben freigeben",
            4: "Exportieren",
        }
        
        self.setup_ui()

    def _apply_window_state(self):
        """Startet das Hauptfenster bevorzugt maximiert (mit Fallback)."""
        try:
            # Windows/Tk: bevorzugter Weg
            self.root.state('zoomed')
        except Exception:
            try:
                # Fallback für Umgebungen, in denen 'zoomed' nicht verfügbar ist
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
    
    def setup_ui(self):
        """Erstellt die Benutzeroberfläche"""
        
        # Hauptframe
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Datei-Auswahl Sektion
        file_frame = ttk.LabelFrame(main_frame, text="Aufgabensammlung laden", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.file_path_var = tk.StringVar()
        ttk.Label(file_frame, text="Word-Datei:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=50).grid(row=0, column=1, padx=(10, 0))
        self.btn_browse = ttk.Button(file_frame, text="Durchsuchen...", command=self.browse_and_load_file)
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

        ttk.Label(file_frame, text="Duplikat-Preset:").grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
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
        ttk.Label(file_frame, textvariable=self.duplicate_threshold_var).grid(
            row=3,
            column=1,
            columnspan=2,
            sticky=tk.W,
            padx=(10, 0),
            pady=(4, 0),
        )
        
        # Kriterien-Auswahl Sektion
        criteria_frame = ttk.LabelFrame(main_frame, text="Auswahlkriterien", padding="10")
        criteria_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Suchbegriffe
        ttk.Label(criteria_frame, text="Suchbegriffe (kommagetrennt):").grid(row=0, column=0, sticky=tk.W)
        self.keywords_var = tk.StringVar()
        ttk.Entry(criteria_frame, textvariable=self.keywords_var, width=60).grid(row=0, column=1, padx=(10, 0))
        
        # Schwierigkeitsgrad
        ttk.Label(criteria_frame, text="Schwierigkeitsgrad:").grid(row=1, column=0, sticky=tk.W, pady=(10, 0))
        self.difficulty_var = tk.StringVar()
        difficulty_combo = ttk.Combobox(criteria_frame, textvariable=self.difficulty_var, 
                                      values=["Alle", "Leicht", "Mittel", "Schwer"])
        difficulty_combo.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=(10, 0))
        difficulty_combo.set("Alle")
        
        # Anzahl der Aufgaben
        ttk.Label(criteria_frame, text="Max. Anzahl Aufgaben:").grid(row=2, column=0, sticky=tk.W, pady=(10, 0))
        self.max_tasks_var = tk.StringVar(value="10")
        ttk.Spinbox(criteria_frame, from_=1, to=100, textvariable=self.max_tasks_var, width=10).grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=(10, 0))
        
        # Wizard-Navigation
        wizard_frame = ttk.LabelFrame(main_frame, text="Import-Assistent", padding="10")
        wizard_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        self.wizard_step_var = tk.StringVar(value="Schritt 1/4: Quelle wählen")
        ttk.Label(wizard_frame, textvariable=self.wizard_step_var).grid(row=0, column=0, sticky=tk.W)

        nav_frame = ttk.Frame(wizard_frame)
        nav_frame.grid(row=0, column=1, sticky=tk.E)
        self.btn_step_prev = ttk.Button(nav_frame, text="← Zurück", command=self._on_prev_step)
        self.btn_step_prev.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_step_next = ttk.Button(nav_frame, text="Weiter →", command=self._on_next_step)
        self.btn_step_next.pack(side=tk.LEFT)

        wizard_frame.columnconfigure(0, weight=1)

        # Aufgaben-Vorschau
        preview_frame = ttk.LabelFrame(main_frame, text="Gefundene Aufgaben", padding="10")
        preview_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        self.wizard_status_var = tk.StringVar(value="Wizard-Status: Keine Datei geladen")
        ttk.Label(preview_frame, textvariable=self.wizard_status_var).grid(
            row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 6)
        )
        
        # Treeview für Aufgabenvorschau
        columns = ("Nr", "Kategorie", "Titel", "Schwierigkeit", "Confidence", "Warnungen", "Suchbegriffe")
        self.task_tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=10, selectmode="extended")
        
        # Spaltenbreiten optimiert setzen
        self.task_tree.heading("Nr", text="Nr")
        self.task_tree.column("Nr", width=50, minwidth=40)
        
        self.task_tree.heading("Kategorie", text="Kategorie")
        self.task_tree.column("Kategorie", width=170, minwidth=120)

        self.task_tree.heading("Titel", text="Titel")
        self.task_tree.column("Titel", width=260, minwidth=180)
        
        self.task_tree.heading("Schwierigkeit", text="Schwierigkeit")
        self.task_tree.column("Schwierigkeit", width=100, minwidth=80)

        self.task_tree.heading("Confidence", text="Confidence")
        self.task_tree.column("Confidence", width=90, minwidth=80)

        self.task_tree.heading("Warnungen", text="Warnungen")
        self.task_tree.column("Warnungen", width=260, minwidth=160)
        
        self.task_tree.heading("Suchbegriffe", text="Suchbegriffe")
        self.task_tree.column("Suchbegriffe", width=200, minwidth=140)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=scrollbar.set)
        
        self.task_tree.tag_configure('confidence_low', background='#ffe6e6')
        self.task_tree.tag_configure('confidence_medium', background='#fff6dd')

        self.task_tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=(10, 0))
        
        self.btn_filter = ttk.Button(button_frame, text="Aufgaben filtern", command=self.filter_tasks)
        self.btn_filter.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_select_all = ttk.Button(button_frame, text="Alle auswählen", command=self.select_all_tasks)
        self.btn_select_all.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_deselect_all = ttk.Button(button_frame, text="Auswahl aufheben", command=self.deselect_all_tasks)
        self.btn_deselect_all.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_approve = ttk.Button(button_frame, text="Auswahl freigeben", command=self.approve_selected_tasks)
        self.btn_approve.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_clear_approvals = ttk.Button(button_frame, text="Freigaben löschen", command=self.clear_task_approvals)
        self.btn_clear_approvals.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_export_selected = ttk.Button(button_frame, text="Markierte exportieren", command=self.export_selected)
        self.btn_export_selected.pack(side=tk.LEFT, padx=(0, 10))
        self.btn_export_all = ttk.Button(button_frame, text="Alle exportieren", command=self.export_all)
        self.btn_export_all.pack(side=tk.LEFT)
        
        # Grid-Konfiguration für responsives Layout
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self._sync_duplicate_mode_ui()
        self._set_step(1, show_message=False)

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
                threshold = f"{float(resolver()):.2f}"
            except Exception:
                threshold = '-'

        if hasattr(self, 'duplicate_threshold_var'):
            self.duplicate_threshold_var.set(
                f"Aktive Duplikat-Schwelle: {threshold} (Modus: {current_mode})"
            )

    def _apply_duplicate_mode_from_ui(self, show_message=True):
        """Übernimmt den im UI gewählten Duplikat-Modus in die Laufzeitregeln."""
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

    def _update_wizard_step_ui(self):
        """Aktualisiert Schrittanzeige, Navigations- und Aktionsbuttons."""
        label = self.step_labels.get(self.current_step, "Unbekannt")
        self.wizard_step_var.set(f"Schritt {self.current_step}/4: {label}")

        max_step = self._max_reachable_step()

        self.btn_step_prev.configure(state=tk.NORMAL if self.current_step > 1 else tk.DISABLED)
        next_enabled = self.current_step < 4 and self.current_step < max_step
        self.btn_step_next.configure(state=tk.NORMAL if next_enabled else tk.DISABLED)

        has_session = self.import_session is not None and max_step >= 2
        enable_workflow = tk.NORMAL if has_session and self.current_step >= 2 else tk.DISABLED

        self.btn_filter.configure(state=enable_workflow)
        self.btn_select_all.configure(state=enable_workflow)
        self.btn_deselect_all.configure(state=enable_workflow)
        self.btn_approve.configure(state=enable_workflow)
        self.btn_clear_approvals.configure(state=enable_workflow)

        export_enabled = tk.NORMAL if has_session and self.current_step >= 4 and max_step >= 4 else tk.DISABLED
        self.btn_export_selected.configure(state=export_enabled)
        self.btn_export_all.configure(state=export_enabled)

    def _refresh_wizard_status(self):
        """Aktualisiert die kompakte Statuszeile für den Wizard-Zustand."""
        if not self.import_session:
            self.wizard_status_var.set("Wizard-Status: Keine Datei geladen")
            self._update_wizard_step_ui()
            return

        stats = self.import_session.get_stats()
        warnings_total = stats.get('warnings', 0)
        low_total = stats.get('low', 0)
        approved = stats.get('approved', 0)
        total = stats.get('total', 0)

        self.wizard_status_var.set(
            f"Wizard-Status: Freigegeben {approved}/{total} | Warnungen {warnings_total} | Low-Confidence {low_total}"
        )
        self._update_wizard_step_ui()
    
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
                            ids_preview = ''
                            if updated_ids:
                                preview_list = updated_ids[:10]
                                ids_preview = "\n\nBetroffene Aufgaben-IDs:\n- " + "\n- ".join(preview_list)
                                if len(updated_ids) > 10:
                                    ids_preview += f"\n- ... (+{len(updated_ids) - 10} weitere)"

                            messagebox.showinfo(
                                "Titel ergänzt",
                                f"Titel wurden für {result.get('changed_tasks', 0)} Aufgabe(n) ergänzt.\n"
                                f"Backup: {result.get('backup_file', '-')}"
                                f"{ids_preview}",
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
            )
            
            self.populate_task_tree(self.loaded_tasks)
            self._refresh_wizard_status()
            stats = self.import_session.get_stats() if self.import_session else {"total": len(self.loaded_tasks)}

            warning_hint = ""
            if diagnostics_report:
                warning_hint = (
                    f"\nWarnungsbehaftete Aufgaben: {diagnostics_report.get('warning_task_count', 0)}"
                    f"\nLow-Confidence: {diagnostics_report.get('low_confidence_count', 0)}"
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
                    "data/Vorlagen/AUFGABEN_MUSTER_STANDARD.docx",
                )
                return

            self._set_step(2, show_message=False)

            messagebox.showinfo("Erfolg", f"{stats['total']} Aufgaben geladen.{warning_hint}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Datei: {str(e)}")

    def import_task_from_word(self):
        """Übernimmt eine neue Aufgabe aus einer separaten Word-Datei in die aktuelle Sammlung."""
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
        max_errors = int(self._rule_value('bulk_max_errors', 5) or 5)
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

    def _ask_import_metadata(self, default_category=None, default_difficulty=None, default_keywords=None, default_title=None):
        """Fragt Metadaten für Aufgabenimport ab und gibt diese normalisiert zurück."""
        rules_default_category = self._rule_value('default_import_metadata.category', '')
        rules_default_difficulty = self._rule_value('default_import_metadata.difficulty', 'mittel')
        rules_default_keywords = self._rule_value('default_import_metadata.keywords', '')
        rules_default_title = self._rule_value('default_import_metadata.title', '')
        allowed_values = self._difficulty_allowed_values()
        difficulty_prompt_values = " | ".join(allowed_values)

        category_default = (
            default_category
            if default_category is not None
            else (self.lek_theme or rules_default_category or "Allgemein")
        )
        difficulty_default = default_difficulty if default_difficulty is not None else rules_default_difficulty
        keywords_default = default_keywords if default_keywords is not None else rules_default_keywords
        title_default = default_title if default_title is not None else rules_default_title
        category = simpledialog.askstring(
            "Kategorie",
            "Kategorie für die neue Aufgabe:",
            initialvalue=category_default,
            parent=self.root,
        )
        if category is None:
            return None

        category = category.strip()
        if not category:
            messagebox.showwarning("Hinweis", "Kategorie darf nicht leer sein.")
            return None

        difficulty = ""
        while True:
            difficulty_input = simpledialog.askstring(
                "Schwierigkeitsgrad",
                f"Schwierigkeitsgrad ({difficulty_prompt_values}):",
                initialvalue=difficulty or difficulty_default,
                parent=self.root,
            )
            if difficulty_input is None:
                return None

            normalized_difficulty = self._normalize_difficulty_value(difficulty_input)
            if normalized_difficulty:
                difficulty = normalized_difficulty
                break

            messagebox.showwarning(
                "Ungültiger Wert",
                f"Bitte genau einen gültigen Schwierigkeitsgrad eingeben: {difficulty_prompt_values}.",
            )

        keywords = simpledialog.askstring(
            "Schlagworte",
            "Schlagworte (kommagetrennt, optional):",
            initialvalue=keywords_default,
            parent=self.root,
        )
        if keywords is None:
            return None

        title = simpledialog.askstring(
            "Titel (optional)",
            "Titel (optional, leer = automatisch aus Quelle ableiten):",
            initialvalue=title_default,
            parent=self.root,
        )
        if title is None:
            return None

        return {
            'category': category,
            'difficulty': difficulty,
            'keywords': keywords,
            'title': str(title or '').strip(),
        }

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

        aliases = self._rule_value('difficulty_rules.aliases', {}) or {}
        default_aliases = {
            'easy': 'leicht',
            'einfach': 'leicht',
            'medium': 'mittel',
            'normal': 'mittel',
            'hard': 'schwer',
            'difficult': 'schwer',
            'komplex': 'schwer',
        }
        merged_aliases = dict(default_aliases)
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
        values = self._rule_value('difficulty_rules.allowed_values', ['leicht', 'mittel', 'schwer']) or []
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
    
    def filter_tasks(self):
        """Filtert Aufgaben basierend auf den angegebenen Kriterien"""
        if not self.loaded_tasks:
            messagebox.showwarning("Warnung", "Bitte laden Sie zuerst eine Aufgabensammlung.")
            return
        
        criteria = {
            'keywords': [kw.strip() for kw in self.keywords_var.get().split(',') if kw.strip()],
            'difficulty': self.difficulty_var.get() if self.difficulty_var.get() != "Alle" else None,
            'max_count': int(self.max_tasks_var.get())
        }
        
        filtered_tasks = self.task_selector.filter_tasks(self.loaded_tasks, criteria)
        self.current_displayed_tasks = filtered_tasks  # Speichere die gefilterten Aufgaben
        self.populate_task_tree(filtered_tasks)
        self._refresh_wizard_status()
        
        messagebox.showinfo("Filter angewendet", f"{len(filtered_tasks)} Aufgaben entsprechen den Kriterien.")
    
    def populate_task_tree(self, tasks):
        """Füllt die Treeview mit den gefundenen Aufgaben"""
        # Alle existierenden Einträge löschen
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        
        # Neue Einträge hinzufügen
        for i, task in enumerate(tasks, 1):
            # Verwende die ursprüngliche Aufgabennummer oder Display-Index
            original_number = task.get('number', i)
            
            # Keywords sicher formatieren
            keywords = task.get('keywords', [])
            keywords_text = ', '.join(keywords) if keywords else ''
            confidence = task.get('confidence', '-')
            warnings = task.get('warnings', [])
            warnings_text = '; '.join(warnings[:2]) if warnings else ''
            
            item_tags = ()
            if confidence == 'low':
                item_tags = ('confidence_low',)
            elif confidence == 'medium':
                item_tags = ('confidence_medium',)

            self.task_tree.insert("", "end", iid=str(original_number), values=(
                original_number,
                task.get('category', 'Ohne Kategorie'),
                task.get('title', 'Ohne Titel'),
                task.get('difficulty', 'Unbekannt'),
                confidence,
                warnings_text,
                keywords_text
            ), tags=item_tags)

    def _selected_task_ids_from_tree(self):
        """Liefert die aktuell markierten Aufgaben-IDs aus der Treeview."""
        selected_ids = []
        for item in self.task_tree.selection():
            values = self.task_tree.item(item).get("values", [])
            if not values:
                continue
            try:
                selected_ids.append(int(values[0]))
            except (TypeError, ValueError):
                continue
        return selected_ids

    def approve_selected_tasks(self):
        """Markiert aktuell ausgewählte Aufgaben als für den Export freigegeben."""
        if not self.import_session:
            messagebox.showwarning("Warnung", "Keine Import-Session vorhanden. Bitte zuerst Datei laden.")
            return

        selected_ids = self._selected_task_ids_from_tree()
        if not selected_ids:
            messagebox.showwarning("Warnung", "Bitte mindestens eine Aufgabe auswählen.")
            return

        for task_id in selected_ids:
            self.import_session.set_task_approval(task_id, True)

        stats = self.import_session.get_stats()
        self._set_step(3, show_message=False)
        self._refresh_wizard_status()
        messagebox.showinfo(
            "Freigabe aktualisiert",
            f"{len(selected_ids)} Aufgabe(n) freigegeben.\n"
            f"Aktuell freigegeben: {stats['approved']} von {stats['total']}.",
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

        # Bei vorhandener Session: Auswahl = Freigabe + Export aus bestätigten Tasks
        if self.import_session:
            for task_id in selected_ids:
                self.import_session.set_task_approval(task_id, True)
            selected_set = set(selected_ids)
            selected_tasks = [
                task for task in self.import_session.get_approved_raw_tasks()
                if task.get('number', 0) in selected_set
            ]
        else:
            # Fallback ohne Session
            current_displayed_tasks = getattr(self, 'current_displayed_tasks', self.loaded_tasks)
            for task_id in selected_ids:
                for task in current_displayed_tasks:
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
        block_difficulty = bool(self._rule_value('difficulty_rules.block_export_on_inconsistent', True))
        allowed_difficulty = {str(v).strip().lower() for v in self._difficulty_allowed_values()}
        block_category = bool(self._rule_value('category_rules.block_export_on_missing', True))
        block_required = bool(self._rule_value('template_rules.block_export_on_missing_required', True))
        missing_values = self._rule_value('category_rules.missing_values', ['', 'ohne kategorie']) or []
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
                        f"#{task.get('number', '?')} {task.get('title', 'Ohne Titel')}"
                    )

            if block_category:
                category_value = str(task.get('category', '') or '').strip().lower()
                category_warning = any('Kategorie fehlt' in str(w) for w in warnings)
                if category_warning or category_value in missing_values_norm:
                    blocked_category_tasks.append(
                        f"#{task.get('number', '?')} {task.get('title', 'Ohne Titel')}"
                    )

            if block_required:
                required_warning = any('Pflichtfeld fehlt:' in str(w) for w in warnings)
                if required_warning:
                    blocked_required_tasks.append(
                        f"#{task.get('number', '?')} {task.get('title', 'Ohne Titel')}"
                    )

        if blocked_difficulty_tasks or blocked_category_tasks or blocked_required_tasks:
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