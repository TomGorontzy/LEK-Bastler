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
import os
import sys
from word_processor import WordProcessor
from task_selector import TaskSelector

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
        
        self.setup_ui()

    def _apply_window_icon(self):
        """Setzt das Fenster-Icon (Titelleiste) wenn verfügbar."""
        try:
            if getattr(sys, 'frozen', False):
                base_dir = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))

            icon_path = os.path.join(base_dir, "app_icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
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
        ttk.Button(file_frame, text="Durchsuchen...", command=self.browse_and_load_file).grid(row=0, column=2, padx=(10, 0))
        
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
        
        # Aufgaben-Vorschau
        preview_frame = ttk.LabelFrame(main_frame, text="Gefundene Aufgaben", padding="10")
        preview_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Treeview für Aufgabenvorschau
        columns = ("Nr", "Kategorie", "Titel", "Schwierigkeit", "Suchbegriffe")
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
        
        self.task_tree.heading("Suchbegriffe", text="Suchbegriffe")
        self.task_tree.column("Suchbegriffe", width=250, minwidth=150)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.task_tree.yview)
        self.task_tree.configure(yscrollcommand=scrollbar.set)
        
        self.task_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(button_frame, text="Aufgaben filtern", command=self.filter_tasks).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Alle auswählen", command=self.select_all_tasks).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Auswahl aufheben", command=self.deselect_all_tasks).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Markierte exportieren", command=self.export_selected).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Alle exportieren", command=self.export_all).pack(side=tk.LEFT)
        
        # Grid-Konfiguration für responsives Layout
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def browse_and_load_file(self):
        """Öffnet einen Dateidialog zur Auswahl der Word-Datei und lädt sie direkt"""
        # Aufgaben-Verzeichnis als Standard setzen
        import sys
        if getattr(sys, 'frozen', False):
            # PyInstaller .exe - verwende Verzeichnis der .exe
            script_dir = os.path.dirname(sys.executable)
        else:
            # Python-Skript - verwende Verzeichnis der .py-Datei
            script_dir = os.path.dirname(os.path.abspath(__file__))
        
        aufgaben_dir = os.path.join(script_dir, "Aufgaben")
        if not os.path.exists(aufgaben_dir):
            aufgaben_dir = script_dir  # Fallback auf Hauptverzeichnis
        
        filename = filedialog.askopenfilename(
            title="Word-Datei auswählen",
            initialdir=aufgaben_dir,
            filetypes=[("Word-Dokumente", "*.docx"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
            self.load_tasks()
    
    def browse_file(self):
        """Öffnet einen Dateidialog zur Auswahl der Word-Datei"""
        # Aufgaben-Verzeichnis als Standard setzen
        import sys
        if getattr(sys, 'frozen', False):
            # PyInstaller .exe - verwende Verzeichnis der .exe
            script_dir = os.path.dirname(sys.executable)
        else:
            # Python-Skript - verwende Verzeichnis der .py-Datei
            script_dir = os.path.dirname(os.path.abspath(__file__))
        
        aufgaben_dir = os.path.join(script_dir, "Aufgaben")
        if not os.path.exists(aufgaben_dir):
            aufgaben_dir = script_dir  # Fallback auf Hauptverzeichnis
        
        filename = filedialog.askopenfilename(
            title="Word-Datei auswählen",
            initialdir=aufgaben_dir,
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
            self.loaded_tasks = self.word_processor.extract_tasks(file_path)
            self.current_displayed_tasks = self.loaded_tasks  # Setze initial alle Aufgaben als angezeigt
            
            # Extrahiere LEK-Thema aus dem Dateinamen
            self.source_filename = os.path.basename(file_path)
            self.lek_theme = self._extract_lek_theme(self.source_filename)
            
            self.populate_task_tree(self.loaded_tasks)
            messagebox.showinfo("Erfolg", f"{len(self.loaded_tasks)} Aufgaben geladen.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Datei: {str(e)}")
    
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
            
            self.task_tree.insert("", "end", values=(
                original_number,
                task.get('category', 'Ohne Kategorie'),
                task.get('title', 'Ohne Titel'),
                task.get('difficulty', 'Unbekannt'),
                keywords_text
            ))
    
    def select_all_tasks(self):
        """Wählt alle Aufgaben in der Treeview aus"""
        for item in self.task_tree.get_children():
            self.task_tree.selection_add(item)
    
    def deselect_all_tasks(self):
        """Hebt die Auswahl aller Aufgaben auf"""
        self.task_tree.selection_remove(self.task_tree.selection())
    
    def export_selected(self):
        """Exportiert nur die markierten Aufgaben in eine neue Word-Datei"""
        selected_items = self.task_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warnung", "Bitte wählen Sie mindestens eine Aufgabe aus.")
            return
        
        # Aufgaben basierend auf der Auswahl sammeln
        current_displayed_tasks = getattr(self, 'current_displayed_tasks', self.loaded_tasks)
        selected_tasks = []
        
        for item in selected_items:
            # Index in der current_displayed_tasks Liste finden
            values = self.task_tree.item(item)["values"]
            original_number = int(values[0])
            
            # Aufgabe mit entsprechender Nummer finden
            for task in current_displayed_tasks:
                if task.get('number', 0) == original_number:
                    selected_tasks.append(task)
                    break
        
        if not selected_tasks:
            messagebox.showwarning("Warnung", "Keine gültigen Aufgaben zur Auswahl gefunden.")
            return
        
        # Export durchführen
        self._perform_export(selected_tasks, f"{len(selected_tasks)} markierte Aufgabe(n)")
    
    def export_all(self):
        """Exportiert alle aktuell angezeigten Aufgaben in eine neue Word-Datei"""
        current_displayed_tasks = getattr(self, 'current_displayed_tasks', self.loaded_tasks)
        
        if not current_displayed_tasks:
            messagebox.showwarning("Warnung", "Keine Aufgaben zum Exportieren vorhanden.")
            return
        
        # Export durchführen
        self._perform_export(current_displayed_tasks, f"alle {len(current_displayed_tasks)} Aufgabe(n)")
    
    def _perform_export(self, tasks_to_export, description):
        """
        Führt den eigentlichen Export durch
        
        Args:
            tasks_to_export: Liste der zu exportierenden Aufgaben
            description: Beschreibung für die Erfolgsmeldung
        """
        # Generiere Standarddateinamen
        from datetime import datetime
        now = datetime.now()
        year_month_day = now.strftime("%Y%m%d")
        time_stamp = now.strftime("%H%M")
        
        # LEK-Export-Verzeichnis erstellen  
        import sys
        if getattr(sys, 'frozen', False):
            # PyInstaller .exe - verwende Verzeichnis der .exe
            script_dir = os.path.dirname(sys.executable)
        else:
            # Python-Skript - verwende Verzeichnis der .py-Datei
            script_dir = os.path.dirname(os.path.abspath(__file__))
        
        leks_dir = os.path.join(script_dir, "LEKs")
        os.makedirs(leks_dir, exist_ok=True)
        
        if self.lek_theme:
            default_filename = f"LEK_{self.lek_theme}_{year_month_day}_{time_stamp}.docx"
        else:
            default_filename = f"LEK_Aufgaben_{year_month_day}_{time_stamp}.docx"
        
        # Speicherpfad wählen
        output_file = filedialog.asksaveasfilename(
            title="Word-Datei speichern",
            defaultextension=".docx",
            initialfile=default_filename,
            initialdir=leks_dir,
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