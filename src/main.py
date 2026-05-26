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
        self._set_step(1, show_message=False)

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
    
    def load_tasks(self):
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
        target_collection = self.file_path_var.get().strip()
        if not target_collection:
            messagebox.showwarning(
                "Hinweis",
                "Bitte zuerst die Ziel-Aufgabensammlung auswählen und laden.",
            )
            return

        if not os.path.exists(target_collection):
            messagebox.showerror("Fehler", "Die ausgewählte Ziel-Aufgabensammlung existiert nicht.")
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

        default_category = self.lek_theme or "Allgemein"
        category = simpledialog.askstring(
            "Kategorie",
            "Kategorie für die neue Aufgabe:",
            initialvalue=default_category,
            parent=self.root,
        )
        if category is None:
            return
        category = category.strip()
        if not category:
            messagebox.showwarning("Hinweis", "Kategorie darf nicht leer sein.")
            return

        difficulty = ""
        while True:
            difficulty_input = simpledialog.askstring(
                "Schwierigkeitsgrad",
                "Schwierigkeitsgrad (leicht | mittel | schwer):",
                initialvalue=difficulty or "mittel",
                parent=self.root,
            )
            if difficulty_input is None:
                return

            normalized_difficulty = self._normalize_difficulty_value(difficulty_input)
            if normalized_difficulty:
                difficulty = normalized_difficulty
                break

            messagebox.showwarning(
                "Ungültiger Wert",
                "Bitte genau einen gültigen Schwierigkeitsgrad eingeben: leicht, mittel oder schwer.",
            )

        keywords = simpledialog.askstring(
            "Schlagworte",
            "Schlagworte (kommagetrennt, optional):",
            initialvalue="",
            parent=self.root,
        )
        if keywords is None:
            return

        try:
            preview = self.word_processor.preview_task_append(
                source_doc_path=source_file,
                target_collection_path=target_collection,
                category=category,
                difficulty=difficulty,
                keywords=keywords,
            )
        except Exception as e:
            messagebox.showerror("Fehler", f"Vorschau fehlgeschlagen: {str(e)}")
            return

        if preview.get('difficulty_inconsistent'):
            messagebox.showwarning(
                "Übernahme blockiert",
                "Der eingegebene Schwierigkeitsgrad ist inkonsistent (mehrere Werte).\n\n"
                "Bitte genau einen Wert verwenden: leicht, mittel oder schwer.",
            )
            return

        lines = preview.get('source_preview_lines', [])
        source_preview_text = "\n".join(f"- {line}" for line in lines) if lines else "- (kein Textvorschau verfügbar)"
        if preview.get('source_paragraph_count', 0) > len(lines):
            source_preview_text += f"\n- ... (+{preview['source_paragraph_count'] - len(lines)} weitere Absätze)"

        details_question = messagebox.askyesno(
            "Details anzeigen",
            "Möchten Sie vor der Übernahme die Detailansicht der Quellinhalte sehen?",
        )
        if details_question:
            self._show_task_import_details(preview)

        proceed = messagebox.askyesno(
            "Übernahme bestätigen",
            "Neue Aufgabe wird in die Aufgabensammlung übernommen.\n\n"
            f"Ziel-ID: {preview.get('next_id', '-')}\n"
            f"Kategorie: {preview.get('category', '-') }\n"
            f"Schwierigkeit: {preview.get('difficulty_normalized', '-') }\n"
            f"Schlagworte: {preview.get('keywords', '-') or '-'}\n"
            f"Quellstruktur: {preview.get('source_paragraph_count', 0)} Absätze, "
            f"{preview.get('source_table_count', 0)} Tabellen\n\n"
            "Quell-Vorschau:\n"
            f"{source_preview_text}\n\n"
            "Fortfahren?",
        )
        if not proceed:
            return

        try:
            result = self.word_processor.append_task_from_document(
                source_doc_path=source_file,
                target_collection_path=target_collection,
                category=category,
                difficulty=difficulty,
                keywords=keywords,
            )
            self.load_tasks()
            messagebox.showinfo(
                "Erfolg",
                "Neue Aufgabe wurde in die Aufgabensammlung übernommen.\n\n"
                f"ID: {result.get('id', '-')}",
            )
        except Exception as e:
            messagebox.showerror("Fehler", f"Übernahme fehlgeschlagen: {str(e)}")

    def _normalize_difficulty_value(self, value):
        """Normalisiert Benutzereingaben auf leicht/mittel/schwer oder liefert None."""
        val = str(value or '').strip().lower()
        if not val:
            return None

        if val in ('leicht', 'mittel', 'schwer'):
            return val

        aliases = {
            'easy': 'leicht',
            'einfach': 'leicht',
            'medium': 'mittel',
            'normal': 'mittel',
            'hard': 'schwer',
            'difficult': 'schwer',
            'komplex': 'schwer',
        }
        return aliases.get(val)

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
        blocked_tasks = []
        for task in tasks_to_export:
            warnings = task.get('warnings', []) or []
            if any('Inkonsistenter Schwierigkeitsgrad' in str(w) for w in warnings):
                blocked_tasks.append(
                    f"#{task.get('number', '?')} {task.get('title', 'Ohne Titel')}"
                )

        if blocked_tasks:
            sample = "\n".join(f"- {entry}" for entry in blocked_tasks[:8])
            if len(blocked_tasks) > 8:
                sample += f"\n- ... (+{len(blocked_tasks) - 8} weitere)"

            messagebox.showwarning(
                "Export blockiert",
                "Der Export wurde gestoppt, weil Inkonsistenzen erkannt wurden.\n\n"
                "Bitte bereinigen Sie die betroffenen Aufgaben direkt in der Quelldatei "
                "und laden Sie diese erneut.\n\n"
                "Betroffene Aufgaben:\n"
                f"{sample}",
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