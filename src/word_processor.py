"""
WordProcessor - Klasse für die Verarbeitung von Word-Dokumenten
"""

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml.shared import qn
from docx.oxml import OxmlElement
from template_manager import TemplateManager
import re
import os
import sys
from pathlib import Path
from copy import deepcopy

class WordProcessor:
    """Klasse für das Lesen und Schreiben von Word-Dokumenten"""
    
    def __init__(self):
        if getattr(sys, 'frozen', False):
            runtime_base = Path(sys.executable).resolve().parent
        else:
            runtime_base = Path(__file__).resolve().parents[1]

        template_folder = runtime_base / 'data' / 'Vorlagen'
        self.template_manager = TemplateManager(str(template_folder))
        self._num_id_map = {}
        self._last_context_report = None
    
    def extract_tasks(self, file_path):
        """
        Extrahiert Aufgaben aus einer Word-Datei.

        Struktur-Regel:
        - Überschrift 1 = Kategorie
        - Überschrift 2 = neue Aufgabe
        - Alle folgenden Elemente bis zur nächsten Überschrift 2 oder 1
          gehören zur aktuellen Aufgabe.
        
        Args:
            file_path (str): Pfad zur Word-Datei
            
        Returns:
            list: Liste von Aufgaben-Dictionaries mit kompletter Struktur
        """
        try:
            doc = Document(file_path)
            tasks = []
            
            all_elements = self._get_all_body_elements(doc)

            current_category = "Ohne Kategorie"
            current_task = None

            def _finalize_current_task():
                if current_task is None:
                    return

                content = []
                for task_element in current_task['all_elements']:
                    if task_element['type'] == 'paragraph' and task_element['content'].text.strip():
                        content.append(task_element['content'].text.strip())
                    elif task_element['type'] == 'table':
                        table_text = self._extract_table_text_for_keywords(task_element['content'])
                        if table_text:
                            content.append(table_text)

                current_task['content'] = content
                current_task['original_paragraphs'] = [
                    elem['content'] for elem in current_task['all_elements'] if elem['type'] == 'paragraph'
                ]
                tasks.append(current_task)

            for element in all_elements:
                if element['type'] != 'paragraph':
                    if current_task is not None:
                        current_task['all_elements'].append(element)
                    continue

                paragraph = element['content']
                paragraph_text = paragraph.text.strip()

                if self._is_heading1(paragraph):
                    _finalize_current_task()
                    current_task = None
                    category_text = paragraph_text
                    current_category = category_text if category_text else "Ohne Kategorie"
                    continue

                if self._is_heading2(paragraph):
                    _finalize_current_task()
                    task_title = paragraph_text or f"Aufgabe {len(tasks) + 1}"
                    current_task = {
                        'number': len(tasks) + 1,
                        'category': current_category,
                        'title': self._extract_task_title(task_title),
                        'content': [],
                        'all_elements': [element],
                        'original_paragraphs': [],
                        'difficulty': 'Mittel',
                        'keywords': []
                    }
                    continue

                if current_task is not None:
                    current_task['all_elements'].append(element)

            _finalize_current_task()

            # Fallback: strukturierte Tabellenaufgaben (z. B. AUFGABEN_GERUEST_WORD)
            if not tasks:
                tasks = self._extract_tasks_from_structured_tables(doc)

            # Metadaten für alle Aufgaben extrahieren
            for task in tasks:
                full_content = ' '.join(task.get('content', []))
                if full_content:
                    explicit_difficulty = self._extract_explicit_difficulty(full_content)
                    if explicit_difficulty:
                        task['difficulty'] = explicit_difficulty
                    else:
                        task['difficulty'] = self._extract_difficulty(full_content)
                    
                    explicit_keywords = self._extract_explicit_keywords(full_content)
                    if explicit_keywords:
                        task['keywords'] = explicit_keywords
                    else:
                        task['keywords'] = self._extract_keywords(full_content)
            
            return tasks
            
        except Exception as e:
            raise Exception(f"Fehler beim Lesen der Word-Datei: {str(e)}")

    def _extract_tasks_from_structured_tables(self, doc):
        """
        Extrahiert Aufgaben aus strukturierten 2-Spalten-Tabellen (Label/Wert).

        Erwartetes Muster je Tabelle (mindestens):
        - ID
        - Aufgabenstellung (Pflicht)
        Optional:
        - Intro/Einleitung
        - Lösungsmöglichkeit/Hinweis
        - Schlagworte
        - Schwierigkeitsgrad
        - Kategorie
        """
        tasks = []
        task_number = 1

        for table in doc.tables:
            parsed_task = self._parse_structured_task_table(table, task_number)
            if parsed_task is None:
                continue

            tasks.append(parsed_task)
            task_number += 1

        return tasks

    def append_task_from_document(
        self,
        source_doc_path,
        target_collection_path,
        category,
        difficulty='Mittel',
        keywords='',
        intro='',
        hint='',
        task_id=''
    ):
        """
        Übernimmt eine neue Aufgabe aus einer Word-Datei in eine tabellenbasierte Aufgabensammlung.

        Die gesamte Quelle (Paragraphen, Tabellen, Formatierungen; best effort inkl. eingebetteter Inhalte)
        wird in das Feld "Aufgabenstellung (Pflicht)" der neuen Tabellen-Aufgabe übernommen.

        Returns:
            dict: Informationen zur übernommenen Aufgabe (id, target_file)
        """
        if not os.path.exists(source_doc_path):
            raise FileNotFoundError(f"Quelldatei nicht gefunden: {source_doc_path}")
        if not os.path.exists(target_collection_path):
            raise FileNotFoundError(f"Zieldatei nicht gefunden: {target_collection_path}")

        source_doc = Document(source_doc_path)
        target_doc = Document(target_collection_path)

        template_table = self._find_first_structured_task_table(target_doc)
        if template_table is None:
            raise ValueError(
                "Die Zieldatei enthält keine erkannte Aufgaben-Tabellenstruktur. "
                "Bitte eine Sammlung auf Basis des Aufgaben-Gerüsts verwenden."
            )

        new_table = self._append_table_clone(target_doc, template_table)

        resolved_id = task_id.strip() if str(task_id or '').strip() else self._generate_next_task_id(target_doc)
        resolved_category = str(category or '').strip() or 'Ohne Kategorie'
        resolved_difficulty = str(difficulty or '').strip() or 'Mittel'

        self._set_structured_table_value(new_table, 'id', resolved_id)
        self._set_structured_table_value(new_table, 'kategorie', resolved_category)
        self._set_structured_table_value(new_table, 'introeinleitungoptional', str(intro or '').strip())
        self._set_structured_table_value(new_table, 'loesungsmoeglichkeithinweisoptional', str(hint or '').strip())
        self._set_structured_table_value(new_table, 'schlagwortekommagetrennt', str(keywords or '').strip())
        self._set_structured_table_value(new_table, 'schwierigkeitsgrad', resolved_difficulty)

        task_cell = self._get_structured_table_value_cell(new_table, 'aufgabenstellungpflicht')
        if task_cell is None:
            task_cell = self._ensure_structured_row(new_table, 'Aufgabenstellung (Pflicht)')

        self._replace_cell_content_from_document(target_doc, task_cell, source_doc)

        target_doc.save(target_collection_path)
        return {
            'id': resolved_id,
            'target_file': target_collection_path,
        }

    def preview_task_append(
        self,
        source_doc_path,
        target_collection_path,
        category,
        difficulty='Mittel',
        keywords='',
    ):
        """
        Liefert eine Vorschau für die Aufgabenübernahme ohne Dateiänderung.

        Returns:
            dict mit Ziel-ID, Quellstruktur, Metadaten und ggf. Validierungshinweisen.
        """
        if not os.path.exists(source_doc_path):
            raise FileNotFoundError(f"Quelldatei nicht gefunden: {source_doc_path}")
        if not os.path.exists(target_collection_path):
            raise FileNotFoundError(f"Zieldatei nicht gefunden: {target_collection_path}")

        source_doc = Document(source_doc_path)
        target_doc = Document(target_collection_path)

        next_id = self._generate_next_task_id(target_doc)
        nonempty_paragraphs = [p.text.strip() for p in source_doc.paragraphs if p.text and p.text.strip()]
        preview_lines = nonempty_paragraphs[:3]
        source_blocks = []

        for child in source_doc._element.body:
            tag_local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag_local == 'p':
                text = ''
                for para in source_doc.paragraphs:
                    if para._element == child:
                        text = para.text.strip()
                        break
                if text:
                    source_blocks.append({'type': 'paragraph', 'text': text})
            elif tag_local == 'tbl':
                table_obj = None
                for table in source_doc.tables:
                    if table._element == child:
                        table_obj = table
                        break

                if table_obj is not None:
                    rows = len(table_obj.rows)
                    cols = len(table_obj.columns) if table_obj.columns else 0
                    first_cell = ''
                    if rows > 0 and cols > 0:
                        first_cell = table_obj.cell(0, 0).text.strip()
                    source_blocks.append({
                        'type': 'table',
                        'rows': rows,
                        'cols': cols,
                        'first_cell': first_cell,
                    })

        normalized_difficulty = self._extract_explicit_difficulty(f"Schwierigkeit: {difficulty}")
        if not normalized_difficulty:
            normalized_difficulty = self._extract_difficulty(str(difficulty or ''))

        return {
            'next_id': next_id,
            'source_paragraph_count': len(nonempty_paragraphs),
            'source_table_count': len(source_doc.tables),
            'source_preview_lines': preview_lines,
            'source_preview_blocks': source_blocks[:10],
            'category': str(category or '').strip() or 'Ohne Kategorie',
            'difficulty_input': str(difficulty or '').strip(),
            'difficulty_normalized': normalized_difficulty,
            'keywords': str(keywords or '').strip(),
            'difficulty_inconsistent': self._has_inconsistent_difficulty(difficulty),
        }

    def _normalize_table_key(self, value):
        """Normalisiert Tabellen-Keys für robuste Label-Erkennung."""
        text = (value or '').strip().lower()
        replacements = {
            'ä': 'ae',
            'ö': 'oe',
            'ü': 'ue',
            'ß': 'ss',
        }
        for src, dst in replacements.items():
            text = text.replace(src, dst)

        text = re.sub(r'\([^\)]*\)', '', text)
        text = re.sub(r'[^a-z0-9]+', '', text)
        return text

    def _find_first_structured_task_table(self, doc):
        """Liefert die erste erkannte Aufgaben-Tabelle (ID + Aufgabenstellung) oder None."""
        for table in doc.tables:
            keys = set()
            for row in table.rows:
                if len(row.cells) < 2:
                    continue
                keys.add(self._normalize_table_key(row.cells[0].text))

            if 'id' in keys and ('aufgabenstellungpflicht' in keys or 'aufgabenstellung' in keys):
                return table
        return None

    def _append_table_clone(self, target_doc, source_table):
        """Kloniert eine Tabelle ans Ende des Dokuments und liefert die neue Tabelle zurück."""
        body = target_doc._element.body
        sect_pr = body.sectPr
        cloned_tbl = deepcopy(source_table._tbl)

        if sect_pr is not None:
            body.insert(list(body).index(sect_pr), cloned_tbl)
        else:
            body.append(cloned_tbl)

        return target_doc.tables[-1]

    def _iter_table_rows_by_norm_key(self, table):
        """Iteriert über Tabellenzeilen als Tupel (norm_key, row)."""
        for row in table.rows:
            if len(row.cells) < 2:
                continue
            key = self._normalize_table_key(row.cells[0].text)
            if key:
                yield key, row

    def _clear_cell_content(self, cell):
        """Entfernt alle Blockinhalte aus einer Tabellenzelle (bis auf tcPr)."""
        tc = cell._tc
        for child in list(tc):
            # tcPr erhalten, Inhaltsknoten löschen
            tag_local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag_local != 'tcPr':
                tc.remove(child)

    def _set_structured_table_value(self, table, key_norm, value):
        """Setzt den Wert (2. Spalte) für einen strukturierten Tabellen-Key."""
        target_cell = self._get_structured_table_value_cell(table, key_norm)
        if target_cell is None:
            label = self._key_norm_to_label(key_norm)
            target_cell = self._ensure_structured_row(table, label)

        self._clear_cell_content(target_cell)
        target_cell.text = str(value or '')

    def _get_structured_table_value_cell(self, table, key_norm):
        """Liefert die Wertzelle (2. Spalte) für key_norm oder None."""
        for row_key, row in self._iter_table_rows_by_norm_key(table):
            if row_key == key_norm:
                return row.cells[1]
        return None

    def _ensure_structured_row(self, table, label):
        """Stellt sicher, dass eine Label-Zeile existiert, und liefert die Wertzelle zurück."""
        wanted = self._normalize_table_key(label)
        for row_key, row in self._iter_table_rows_by_norm_key(table):
            if row_key == wanted:
                return row.cells[1]

        new_row = table.add_row()
        new_row.cells[0].text = label
        new_row.cells[1].text = ''
        return new_row.cells[1]

    def _key_norm_to_label(self, key_norm):
        """Mappt normalisierte Keys auf lesbare Standard-Labels."""
        mapping = {
            'id': 'ID',
            'kategorie': 'Kategorie',
            'introeinleitungoptional': 'Intro/Einleitung (optional)',
            'aufgabenstellungpflicht': 'Aufgabenstellung (Pflicht)',
            'loesungsmoeglichkeithinweisoptional': 'Lösungsmöglichkeit/Hinweis (optional)',
            'schlagwortekommagetrennt': 'Schlagworte (kommagetrennt)',
            'schwierigkeitsgrad': 'Schwierigkeitsgrad',
        }
        return mapping.get(key_norm, key_norm)

    def _replace_cell_content_from_document(self, target_doc, target_cell, source_doc):
        """
        Ersetzt den Zelleninhalt durch den Inhalt eines Quell-Dokuments.

        Hinweis: Beziehungen werden best effort in das Zieldokument übernommen.
        """
        self._clear_cell_content(target_cell)
        tc = target_cell._tc

        # Sicherstellen, dass Beziehungen pro old_rId nur einmal neu angelegt werden
        rel_map = {}

        for child in source_doc._element.body:
            tag_local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag_local not in ('p', 'tbl'):
                continue

            cloned = deepcopy(child)
            self._remap_relationship_ids_in_xml(
                xml_element=cloned,
                source_part=source_doc.part,
                target_part=target_doc.part,
                rel_map=rel_map,
            )
            tc.append(cloned)

        # Word erwartet mindestens einen Absatz in der Zelle
        has_block = any((c.tag.split('}')[-1] if '}' in c.tag else c.tag) in ('p', 'tbl') for c in tc)
        if not has_block:
            target_cell.text = ''

    def _remap_relationship_ids_in_xml(self, xml_element, source_part, target_part, rel_map):
        """
        Remappt r:* Beziehungs-IDs in einem XML-Element von source_part auf target_part.
        """
        rel_ns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

        for node in xml_element.iter():
            for attr_name, old_rid in list(node.attrib.items()):
                if not attr_name.startswith(rel_ns):
                    continue

                if old_rid in rel_map:
                    node.set(attr_name, rel_map[old_rid])
                    continue

                new_rid = self._clone_relationship(source_part, target_part, old_rid)
                if new_rid:
                    rel_map[old_rid] = new_rid
                    node.set(attr_name, new_rid)

    def _clone_relationship(self, source_part, target_part, old_rid):
        """
        Klont eine Relationship von source_part nach target_part und liefert neue rId.

        Best effort: Bei nicht unterstützten Relationship-Typen wird alte rId zurückgegeben.
        """
        try:
            rel = source_part.rels[old_rid]
        except Exception:
            return old_rid

        try:
            if getattr(rel, 'is_external', False):
                return target_part.relate_to(rel.target_ref, rel.reltype, is_external=True)

            target_obj = getattr(rel, 'target_part', None) or getattr(rel, '_target', None)
            if target_obj is None:
                return old_rid

            return target_part.relate_to(target_obj, rel.reltype)
        except Exception:
            return old_rid

    def _generate_next_task_id(self, target_doc):
        """Erzeugt eine nächste Aufgaben-ID auf Basis vorhandener IDs (A-001, A-001.1, ...)."""
        ids = []

        for table in target_doc.tables:
            id_cell = self._get_structured_table_value_cell(table, 'id')
            if id_cell is None:
                continue
            val = id_cell.text.strip()
            if val:
                ids.append(val)

        if not ids:
            return 'A-001'

        nums = []
        for val in ids:
            m = re.search(r'(\d+)(?!.*\d)', val)
            if m:
                try:
                    nums.append(int(m.group(1)))
                except ValueError:
                    pass

        if not nums:
            return 'A-001'

        return f"A-{max(nums) + 1:03d}"

    def _parse_structured_task_table(self, table, task_number):
        """Parst eine einzelne strukturierte Aufgaben-Tabelle in ein Task-Dictionary."""
        if not table.rows or len(table.columns) < 2:
            return None

        values_by_key = {}

        for row in table.rows:
            if len(row.cells) < 2:
                continue

            raw_key = row.cells[0].text.strip()
            raw_value = row.cells[1].text.strip()
            norm_key = self._normalize_table_key(raw_key)
            if not norm_key:
                continue

            values_by_key[norm_key] = raw_value

        task_text = values_by_key.get('aufgabenstellungpflicht') or values_by_key.get('aufgabenstellung')
        if not task_text:
            return None

        intro_text = values_by_key.get('introeinleitungoptional') or values_by_key.get('einleitung') or ''
        hint_text = values_by_key.get('loesungsmoeglichkeithinweisoptional') or values_by_key.get('hinweis') or ''
        difficulty_raw = values_by_key.get('schwierigkeitsgrad') or values_by_key.get('schwierigkeit') or ''
        keywords_raw = values_by_key.get('schlagwortekommagetrennt') or values_by_key.get('schlagworte') or ''
        category = values_by_key.get('kategorie') or 'Ohne Kategorie'
        pre_warnings = []

        if self._has_inconsistent_difficulty(difficulty_raw):
            pre_warnings.append(
                "Inkonsistenter Schwierigkeitsgrad erkannt (mehrere Werte). Bitte vor Export in der Quelle bereinigen."
            )

        if difficulty_raw:
            explicit_difficulty = self._extract_explicit_difficulty(f"Schwierigkeit: {difficulty_raw}")
            difficulty = explicit_difficulty if explicit_difficulty else self._extract_difficulty(difficulty_raw)
        else:
            difficulty = self._extract_difficulty(task_text)

        if keywords_raw:
            keywords = [k.strip().lower() for k in keywords_raw.split(',') if k.strip()]
        else:
            keywords = self._extract_keywords(task_text)

        content_lines = []
        if intro_text:
            content_lines.append(f"Einleitung: {intro_text}")
        content_lines.append(task_text)
        if hint_text:
            content_lines.append(f"Hinweis: {hint_text}")
        if difficulty_raw:
            content_lines.append(f"Schwierigkeit: {difficulty_raw}")
        if keywords_raw:
            content_lines.append(f"Schlüsselwörter: {keywords_raw}")

        task_id = values_by_key.get('id') or str(task_number)
        title_source = task_text.split('\n', 1)[0].strip()

        return {
            'number': task_number,
            'id': task_id,
            'category': category,
            'title': self._extract_task_title(title_source),
            'content': content_lines,
            'all_elements': [{'type': 'table', 'content': table}],
            'original_paragraphs': [],
            'difficulty': difficulty,
            'keywords': keywords,
            'pre_warnings': pre_warnings,
        }

    def extract_tasks_with_diagnostics(self, file_path):
        """
        Extrahiert Aufgaben und ergänzt Diagnoseinformationen pro Aufgabe.

        Returns:
            tuple[list, dict]: (tasks, diagnostics_report)
        """
        tasks = self.extract_tasks(file_path)

        task_diagnostics = []
        low_confidence_count = 0
        warning_count = 0

        for task in tasks:
            diagnostic = self._build_task_diagnostic(task)
            task['intro'] = diagnostic['intro']
            task['warnings'] = diagnostic['warnings']
            task['confidence'] = diagnostic['confidence']

            if diagnostic['confidence'] == 'low':
                low_confidence_count += 1
            if diagnostic['warnings']:
                warning_count += 1

            task_diagnostics.append({
                'number': task.get('number', 0),
                'title': task.get('title', ''),
                'confidence': diagnostic['confidence'],
                'warnings': list(diagnostic['warnings']),
            })

        global_warnings = []
        if len(tasks) == 0:
            global_warnings.append("Keine Aufgaben erkannt. Prüfen Sie die Formatierung (Überschrift 2 für Aufgaben).")

        report = {
            'task_count': len(tasks),
            'low_confidence_count': low_confidence_count,
            'warning_task_count': warning_count,
            'global_warnings': global_warnings,
            'task_diagnostics': task_diagnostics,
        }

        return tasks, report

    def _build_task_diagnostic(self, task):
        """
        Erstellt Diagnosemetadaten (intro/warnings/confidence) für eine Aufgabe.
        """
        title = (task.get('title') or '').strip()
        content_lines = [str(line).strip() for line in (task.get('content') or []) if str(line).strip()]
        keywords = task.get('keywords') or []
        difficulty = (task.get('difficulty') or '').strip()

        warnings = list(task.get('pre_warnings') or [])

        if not content_lines:
            warnings.append('Aufgabe enthält keinen verwertbaren Inhalt.')

        if title and len(title) < 5:
            warnings.append('Aufgabentitel ist sehr kurz.')

        if not keywords:
            warnings.append('Keine Schlüsselwörter erkannt.')

        if difficulty in ('', 'Unbekannt'):
            warnings.append('Schwierigkeit ist nicht eindeutig.')

        intro_lines = self._extract_intro_lines(content_lines)
        if intro_lines:
            warnings.append('Einleitungs-/Kontexttext erkannt (Bitte prüfen).')

        # Confidence-Heuristik (Sprint 1)
        score = 100
        for warning in warnings:
            if 'keinen verwertbaren Inhalt' in warning:
                score -= 45
            elif 'Einleitungs-/Kontexttext' in warning:
                score -= 20
            else:
                score -= 10

        if score >= 80:
            confidence = 'high'
        elif score >= 50:
            confidence = 'medium'
        else:
            confidence = 'low'

        return {
            'intro': intro_lines,
            'warnings': warnings,
            'confidence': confidence,
        }

    def _extract_intro_lines(self, content_lines):
        """
        Erkennt mögliche Intro-/Kontext-Zeilen am Aufgabenanfang.
        """
        if not content_lines:
            return []

        intro_markers = (
            'einleitung',
            'intro',
            'hinweis',
            'kontext',
            'ausgangslage',
            'vorbemerkung',
        )

        intro = []
        for line in content_lines[:2]:  # nur die ersten Zeilen als Intro-Kandidaten
            line_lower = line.lower()
            if any(marker in line_lower for marker in intro_markers):
                intro.append(line)

        return intro
    
    def _get_heading_level(self, style_name):
        """
        Bestimmt das korrekte Heading-Level basierend auf dem Style-Namen
        
        Args:
            style_name: Name des Word-Styles
            
        Returns:
            int: Das entsprechende Heading-Level (1-9)
        """
        if not style_name:
            return 1
            
        style_lower = style_name.lower()
        
        # Direkte Zuordnung für deutsche und englische Styles
        heading_map = {
            'heading 1': 1, 'überschrift 1': 1,
            'heading 2': 2, 'überschrift 2': 2,
            'heading 3': 3, 'überschrift 3': 3,
            'heading 4': 4, 'überschrift 4': 4,
            'heading 5': 5, 'überschrift 5': 5,
            'heading 6': 6, 'überschrift 6': 6,
            'heading 7': 7, 'überschrift 7': 7,
            'heading 8': 8, 'überschrift 8': 8,
            'heading 9': 9, 'überschrift 9': 9
        }
        
        if style_lower in heading_map:
            return heading_map[style_lower]
        
        # Fallback: Versuche Zahl aus Style-Namen zu extrahieren
        import re
        match = re.search(r'\d+', style_name)
        if match:
            level = int(match.group())
            return min(max(level, 1), 9)  # Begrenze auf 1-9
        
        return 1  # Standard-Fallback

    def _is_heading1(self, paragraph):
        """
        Überprüft, ob ein Paragraph eine Überschrift 1 ist
        
        Args:
            paragraph: Word-Paragraph-Objekt
            
        Returns:
            bool: True wenn es eine Überschrift 1 ist
        """
        try:
            if paragraph.style:
                style_name = paragraph.style.name
                return (style_name == 'Heading 1' or 
                        style_name == 'Überschrift 1' or
                        style_name.lower() == 'heading 1' or
                        style_name.lower() == 'überschrift 1')
        except:
            # Fallback: Prüfe auf Textmuster wenn Style nicht verfügbar
            text = paragraph.text.strip()
            if text:
                return self._is_task_start(text, paragraph)
        
        return False

    def _is_heading2(self, paragraph):
        """
        Überprüft, ob ein Paragraph eine Überschrift 2 ist.

        Args:
            paragraph: Word-Paragraph-Objekt

        Returns:
            bool: True wenn es eine Überschrift 2 ist
        """
        try:
            if paragraph.style:
                style_name = paragraph.style.name
                return (
                    style_name == 'Heading 2'
                    or style_name == 'Überschrift 2'
                    or style_name.lower() == 'heading 2'
                    or style_name.lower() == 'überschrift 2'
                )
        except:
            return False

        return False

    def _get_all_body_elements(self, doc):
        """
        Sammelt alle Body-Elemente (Paragraphen und Tabellen) in der richtigen Reihenfolge
        
        Args:
            doc: Word-Dokument
            
        Returns:
            list: Liste von Element-Dictionaries mit 'type' und 'content'
        """
        elements = []
        
        # Nutze XML-Struktur für korrekte Reihenfolge
        body = doc._element.body
        
        for child in body:
            tag_name = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            
            if tag_name == 'p':  # Paragraph
                # Finde entsprechendes Paragraph-Objekt
                for para in doc.paragraphs:
                    if para._element == child:
                        elements.append({
                            'type': 'paragraph',
                            'content': para
                        })
                        break
                        
            elif tag_name == 'tbl':  # Table
                # Finde entsprechendes Table-Objekt
                for table in doc.tables:
                    if table._element == child:
                        elements.append({
                            'type': 'table',
                            'content': table
                        })
                        break

            elif tag_name == 'sdt':  # Inhaltssteuerelement (Content Control)
                elements.append({
                    'type': 'sdt',
                    'content': child  # Kein python-docx-Wrapper – raw XML-Element
                })
        
        return elements

    def _extract_table_text_for_keywords(self, table):
        """
        Extrahiert Text aus Tabelle für Keyword-Analyse
        
        Args:
            table: Word-Table-Objekt
            
        Returns:
            str: Tabellen-Text zusammengefasst
        """
        try:
            table_text = []
            for row in table.rows:
                row_text = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    if cell_text:
                        row_text.append(cell_text)
                if row_text:
                    table_text.append(' '.join(row_text))
            return ' '.join(table_text)
        except:
            return ""
    
    def _create_cover_page(self, doc, lek_theme=""):
        """
        Erstellt ein Deckblatt (Semaphor) mit Farbschema basierend auf #00a499
        
        Args:
            doc: Das Word-Dokument
            lek_theme (str): LEK-Thema für das Deckblatt
        """
        # Farbpalette basierend auf #00a499
        primary_color = RGBColor(0, 164, 153)      # #00a499
        light_color = RGBColor(102, 204, 195)      # Hellere Version
        dark_color = RGBColor(0, 120, 112)         # Dunklere Version
        
        # Haupttitel
        title_paragraph = doc.add_paragraph()
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_paragraph.add_run("LERNERFOLGSKONTROLLE")
        title_run.font.name = 'Aptos'
        title_run.font.size = Pt(24)
        title_run.font.color.rgb = primary_color
        title_run.bold = True
        
        # Leerzeilen
        for _ in range(3):
            doc.add_paragraph()
        
        # Semaphor-Design (vereinfacht mit Text-Elementen)
        semaphor_paragraph = doc.add_paragraph()
        semaphor_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Oberer Balken
        top_run = semaphor_paragraph.add_run("━━━━━━━━━━━━━━━━━━━━")
        top_run.font.size = Pt(16)
        top_run.font.color.rgb = dark_color
        top_run.bold = True
        
        doc.add_paragraph()
        
        # Mittlerer Bereich
        middle_paragraph = doc.add_paragraph()
        middle_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if lek_theme:
            middle_run = middle_paragraph.add_run(f"◆  {lek_theme.upper()}  ◆")
        else:
            middle_run = middle_paragraph.add_run("◆  AUFGABENSAMMLUNG  ◆")
            
        middle_run.font.name = 'Aptos'
        middle_run.font.size = Pt(18)
        middle_run.font.color.rgb = light_color
        middle_run.bold = True
        
        doc.add_paragraph()
        
        # Unterer Balken
        bottom_paragraph = doc.add_paragraph()
        bottom_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        bottom_run = bottom_paragraph.add_run("━━━━━━━━━━━━━━━━━━━━")
        bottom_run.font.size = Pt(16)
        bottom_run.font.color.rgb = dark_color
        bottom_run.bold = True
        
        # Weitere Leerzeilen
        for _ in range(8):
            doc.add_paragraph()
        
        # Datum
        date_paragraph = doc.add_paragraph()
        date_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        from datetime import datetime
        date_run = date_paragraph.add_run(f"Erstellt am {datetime.now().strftime('%d.%m.%Y')}")
        date_run.font.name = 'Aptos'
        date_run.font.size = Pt(12)
        date_run.font.color.rgb = primary_color
        date_run.italic = True
    
    def _copy_elements_with_formatting(self, doc, elements):
        """
        Kopiert alle Elemente (Paragraphen und Tabellen) mit ihrer ursprünglichen Formatierung
        
        Args:
            doc: Ziel-Word-Dokument
            elements: Liste von Element-Dictionaries mit 'type' und 'content'
        """
        for element in elements:
            if element['type'] == 'paragraph':
                # Verwende bestehende Paragraph-Kopierfunktion
                self._copy_paragraphs_with_formatting(doc, [element['content']])
                
            elif element['type'] == 'table':
                # Kopiere Tabelle mit vollständiger Formatierung
                self._copy_table_with_formatting(doc, element['content'])

    def _copy_elements_for_lek(self, doc, elements):
        """
        Spezielle Kopierfunktion für LEK-Erstellung mit angepassten Überschrift-Levels
        
        Args:
            doc: Ziel-Word-Dokument (LEK)
            elements: Liste von Element-Dictionaries mit 'type' und 'content'
        """
        from copy import deepcopy
        from docx.oxml.ns import qn as _qn

        body = doc._element.body
        sect_pr = body.find(_qn('w:sectPr'))

        for element in elements:
            try:
                etype = element['type']
                if etype in ('paragraph', 'table'):
                    xml_elem = element['content']._element
                elif etype == 'sdt':
                    xml_elem = element['content']
                else:
                    continue

                new_elem = deepcopy(xml_elem)
                self._remap_num_ids_in_element(new_elem, self._num_id_map)
                if sect_pr is not None:
                    body.insert(list(body).index(sect_pr), new_elem)
                else:
                    body.append(new_elem)
            except Exception as e:
                # Fallback: Fehlermeldung als Absatz einfügen
                doc.add_paragraph(f"[Element konnte nicht kopiert werden: {str(e)[:80]}]")

    def _copy_paragraphs_for_lek(self, doc, paragraphs):
        """
        Kopiert Paragraphen für LEK-Erstellung mit angepassten Überschrift-Levels
        
        Args:
            doc: Ziel-Word-Dokument (LEK)
            paragraphs: Liste von Quell-Paragraphen
        """
        for paragraph in paragraphs:
            if not paragraph.text.strip():
                # Leere Absätze als solche übernehmen
                doc.add_paragraph()
                continue
            
            # Prüfe ob es eine Überschrift ist
            if (paragraph.style and 
                (paragraph.style.name.startswith('Heading') or 
                 paragraph.style.name.startswith('Überschrift') or
                 'Heading' in paragraph.style.name or
                 'Überschrift' in paragraph.style.name)):
                
                # Extrahiere ursprüngliches Level und behalte es bei (KEIN Shift zu Level 2!)
                original_level = self._get_heading_level(paragraph.style.name)
                
                # Erstelle Heading mit ORIGINALEM Level
                new_paragraph = doc.add_heading(paragraph.text, level=original_level)
                
            else:
                # Normaler Absatz - kopiere mit bestehender Logik
                new_paragraph = doc.add_paragraph()
                
                # Absatzformatierung kopieren (vereinfacht für LEK)
                try:
                    if paragraph.style and not paragraph.style.name.startswith('Heading') and not paragraph.style.name.startswith('Überschrift'):
                        try:
                            new_paragraph.style = paragraph.style
                        except:
                            pass
                except:
                    pass
                
                # Runs mit Formatierung kopieren
                for run in paragraph.runs:
                    new_run = new_paragraph.add_run(run.text)
                    try:
                        if run.font:
                            if run.font.name:
                                new_run.font.name = run.font.name
                            if run.font.size:
                                new_run.font.size = run.font.size
                            new_run.bold = run.bold
                            new_run.italic = run.italic
                            new_run.underline = run.underline
                    except:
                        pass

    def _copy_table_with_formatting(self, doc, source_table):
        """
        Kopiert eine Tabelle mit vollständiger Formatierung
        
        Args:
            doc: Ziel-Word-Dokument
            source_table: Quell-Tabelle
        """
        try:
            # Erstelle neue Tabelle mit gleicher Struktur
            rows = len(source_table.rows)
            cols = len(source_table.columns) if source_table.columns else 1
            
            new_table = doc.add_table(rows=rows, cols=cols)
            
            # Kopiere Zellinhalt und Formatierung
            for row_idx, source_row in enumerate(source_table.rows):
                target_row = new_table.rows[row_idx]
                
                for cell_idx, source_cell in enumerate(source_row.cells):
                    if cell_idx < len(target_row.cells):
                        target_cell = target_row.cells[cell_idx]
                        
                        # Kopiere Text-Inhalt
                        target_cell.text = source_cell.text
                        
                        # Kopiere Zellenformatierung (Basic)
                        try:
                            # Paragraph-Formatierung in der Zelle übernehmen
                            if source_cell.paragraphs:
                                source_para = source_cell.paragraphs[0]
                                target_para = target_cell.paragraphs[0]
                                
                                # Alignment kopieren
                                if source_para.paragraph_format.alignment:
                                    target_para.paragraph_format.alignment = source_para.paragraph_format.alignment
                                
                                # Text-Formatierung kopieren
                                if source_para.runs:
                                    # Lösche bestehenden Text
                                    target_para.clear()
                                    for run in source_para.runs:
                                        new_run = target_para.add_run(run.text)
                                        if run.bold:
                                            new_run.bold = True
                                        if run.italic:
                                            new_run.italic = True
                                        if run.underline:
                                            new_run.underline = True
                                        if run.font.size:
                                            new_run.font.size = run.font.size
                        except:
                            # Formatierung fehlgeschlagen, aber Text ist da
                            pass
            
        except Exception as e:
            # Fallback: Einfache Tabellen-Ersetzung
            doc.add_paragraph(f"[Tabelle mit {len(source_table.rows)} Zeilen konnte nicht vollständig kopiert werden]")

    def _copy_paragraphs_with_formatting(self, doc, paragraphs):
        """
        Kopiert Paragraphen mit ihrer ursprünglichen Formatierung
        
        Args:
            doc: Ziel-Word-Dokument
            paragraphs: Liste von Quell-Paragraphen
        """
        for paragraph in paragraphs:
            if not paragraph.text.strip():
                # Leere Absätze als solche übernehmen
                doc.add_paragraph()
                continue
            
            # Neuen Absatz erstellen - prüfe zuerst ob es eine Überschrift ist
            if (paragraph.style and 
                (paragraph.style.name.startswith('Heading') or 
                 paragraph.style.name.startswith('Überschrift') or
                 'Heading' in paragraph.style.name or
                 'Überschrift' in paragraph.style.name)):
                
                # Extrahiere Heading-Level präzise basierend auf Style-Namen
                level = self._get_heading_level(paragraph.style.name)
                
                # Erstelle Heading mit Text
                new_paragraph = doc.add_heading(paragraph.text, level=level)
                
            else:
                # Normaler Absatz
                new_paragraph = doc.add_paragraph()
            
            # Absatzformatierung kopieren
            try:
                if paragraph.style:
                    # Versuche den Style zu übernehmen, außer bei Überschriften
                    if not paragraph.style.name.startswith('Heading') and not paragraph.style.name.startswith('Überschrift'):
                        try:
                            new_paragraph.style = paragraph.style
                        except:
                            # Falls Style nicht verfügbar, verwende Normal
                            pass
                
                # Absatz-Formatierung übernehmen
                if paragraph.paragraph_format:
                    try:
                        new_paragraph.paragraph_format.alignment = paragraph.paragraph_format.alignment
                        new_paragraph.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
                        new_paragraph.paragraph_format.right_indent = paragraph.paragraph_format.right_indent
                        new_paragraph.paragraph_format.first_line_indent = paragraph.paragraph_format.first_line_indent
                        new_paragraph.paragraph_format.space_before = paragraph.paragraph_format.space_before
                        new_paragraph.paragraph_format.space_after = paragraph.paragraph_format.space_after
                        new_paragraph.paragraph_format.line_spacing = paragraph.paragraph_format.line_spacing
                        new_paragraph.paragraph_format.line_spacing_rule = paragraph.paragraph_format.line_spacing_rule
                    except:
                        # Falls Formatierung nicht kopiert werden kann, ignorieren
                        pass
            except:
                # Falls Style-Zugriff fehlschlägt, ignorieren
                pass
            
            # Runs mit Formatierung kopieren (nur wenn es kein Heading ist, da dort Text schon gesetzt wurde)
            if not (paragraph.style and 
                    (paragraph.style.name.startswith('Heading') or 
                     paragraph.style.name.startswith('Überschrift') or
                     'Heading' in paragraph.style.name or
                     'Überschrift' in paragraph.style.name)):
                
                for run in paragraph.runs:
                    new_run = new_paragraph.add_run(run.text)
                    
                    try:
                        # Font-Formatierung kopieren
                        if run.font:
                            if run.font.name:
                                new_run.font.name = run.font.name
                            if run.font.size:
                                new_run.font.size = run.font.size
                            if run.font.color and run.font.color.rgb:
                                new_run.font.color.rgb = run.font.color.rgb
                            
                            new_run.bold = run.bold
                            new_run.italic = run.italic
                            new_run.underline = run.underline
                            
                            # Weitere Font-Eigenschaften
                            try:
                                new_run.font.subscript = run.font.subscript
                                new_run.font.superscript = run.font.superscript
                                new_run.font.strike = run.font.strike
                                new_run.font.small_caps = run.font.small_caps
                                new_run.font.all_caps = run.font.all_caps
                            except:
                                # Ignoriere erweiterte Formatierungen falls nicht verfügbar
                                pass
                    except:
                        # Falls Font-Formatierung nicht kopiert werden kann, ignorieren
                        pass
    
    def _is_task_start(self, text, paragraph):
        """
        Erkennt, ob ein Text der Beginn einer neuen Aufgabe ist
        Priorisiert Überschrift 1 Format, fallback auf Textmuster
        
        Args:
            text (str): Zu prüfender Text
            paragraph: Word-Paragraph-Objekt
            
        Returns:
            bool: True wenn Aufgabenbeginn erkannt
        """
        # Primäre Erkennung: Überschrift 1 Formatvorlage
        try:
            if paragraph.style.name == 'Heading 1' or paragraph.style.name == 'Überschrift 1':
                return True
        except:
            # Falls Style-Zugriff fehlschlägt, weiter mit Textmustern
            pass
        
        # Sekundäre Erkennung: Verschiedene Textmuster für Aufgabenbeginne
        return self._matches_task_start_text(text)

    def _matches_task_start_text(self, text):
        """
        Prüft rein textbasiert, ob ein String wie ein Aufgabenstart aussieht.
        """
        patterns = [
            r'^Aufgabe\s*\d+',
            r'^\d+\.',
            r'^Übung\s*\d+',
            r'^Frage\s*\d+',
            r'^Problem\s*\d+',
            r'^Nr\.\s*\d+',
            r'^\w\)',  # a), b), etc.
        ]
        
        for pattern in patterns:
            if re.match(pattern, text, re.IGNORECASE):
                return True

        return False

    def _extract_task_title(self, text):
        """
        Extrahiert den Titel einer Aufgabe
        
        Args:
            text (str): Text der ersten Zeile der Aufgabe
            
        Returns:
            str: Extrahierter Titel
        """
        # Entfernt Nummerierung und behält den Rest
        cleaned = re.sub(r'^(Aufgabe\s*\d+[:\.]?\s*|^\d+\.\s*|^Übung\s*\d+[:\.]?\s*)', '', text, flags=re.IGNORECASE)
        
        # Falls der Text sehr lang ist, nur die ersten 50 Zeichen nehmen
        if len(cleaned) > 50:
            cleaned = cleaned[:47] + "..."
        
        return cleaned.strip() or "Ohne Titel"
    
    def _extract_difficulty(self, text):
        """
        Versucht den Schwierigkeitsgrad aus dem Text zu extrahieren
        
        Args:
            text (str): Zu analysierender Text
            
        Returns:
            str: Schwierigkeitsgrad (Leicht, Mittel, Schwer)
        """
        text_lower = text.lower()
        
        if any(word in text_lower for word in ['leicht', 'einfach', 'basic', 'anfänger']):
            return 'Leicht'
        elif any(word in text_lower for word in ['schwer', 'difficult', 'komplex', 'advanced']):
            return 'Schwer'
        elif any(word in text_lower for word in ['mittel', 'medium', 'intermediate']):
            return 'Mittel'
        
        # Standard-Schwierigkeitsgrad basierend auf Textlänge
        if len(text) > 200:
            return 'Schwer'
        elif len(text) > 100:
            return 'Mittel'
        else:
            return 'Leicht'
    
    def _extract_keywords(self, text):
        """
        Extrahiert relevante Schlüsselwörter aus dem Text
        
        Args:
            text (str): Zu analysierender Text
            
        Returns:
            list: Liste von Schlüsselwörtern
        """
        # Stopwörter, die ignoriert werden sollen
        stop_words = {
            'der', 'die', 'das', 'ein', 'eine', 'und', 'oder', 'aber', 'von', 'zu', 'mit', 'bei', 'für',
            'ist', 'sind', 'war', 'waren', 'hat', 'haben', 'wird', 'werden', 'kann', 'können', 'soll',
            'sollte', 'auf', 'in', 'an', 'über', 'unter', 'durch', 'gegen', 'ohne', 'um', 'nach', 'vor'
        }
        
        # Text in Wörter aufteilen und bereinigen
        words = re.findall(r'\b\w+\b', text.lower())
        
        # Filtern: mindestens 3 Zeichen, keine Stopwörter, keine Zahlen
        keywords = [
            word for word in words 
            if len(word) >= 3 
            and word not in stop_words 
            and not word.isdigit()
        ]
        
        # Nur die häufigsten/relevantesten Keywords zurückgeben
        return list(set(keywords))[:10]
    
    def _extract_explicit_keywords(self, text):
        """
        Extrahiert explizite Schlüsselwörter aus "Schlüsselwörter:" Zeilen
        
        Args:
            text (str): Zu analysierender Text
            
        Returns:
            list: Liste von explizit angegebenen Schlüsselwörtern oder leere Liste
        """
        import re
        
        # Suche nach "Schlüsselwörter:" gefolgt von kommagetrenner Liste
        patterns = [
            r'Schlüsselwörter:\s*([^\n\r]+)',
            r'Schlüsselwort:\s*([^\n\r]+)',
            r'Keywords:\s*([^\n\r]+)',
            r'Stichwörter:\s*([^\n\r]+)',
            r'Stichworte:\s*([^\n\r]+)'
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                # Nehme den ersten/längsten Treffer
                keywords_text = max(matches, key=len) if len(matches) > 1 else matches[0]
                
                # Trenne an Kommas und bereinige
                keywords = []
                for keyword in keywords_text.split(','):
                    cleaned = keyword.strip()
                    if cleaned and len(cleaned) >= 2:  # Mindestens 2 Zeichen
                        keywords.append(cleaned.lower())
                
                if keywords:
                    return keywords[:10]  # Maximal 10 Keywords
        
        return []  # Keine expliziten Keywords gefunden
    
    def _extract_explicit_difficulty(self, text):
        """
        Extrahiert explizite Schwierigkeitsangaben aus "Schwierigkeit:" Zeilen
        
        Args:
            text (str): Zu analysierender Text
            
        Returns:
            str: Explizit angegebene Schwierigkeit oder None
        """
        import re
        
        # Suche nach "Schwierigkeit:" gefolgt von Schwierigkeitsgrad
        patterns = [
            r'Schwierigkeit:\s*([^\n\r]+)',
            r'Schwierigkeitsgrad:\s*([^\n\r]+)',
            r'Difficulty:\s*([^\n\r]+)',
            r'Level:\s*([^\n\r]+)'
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                difficulty_text = matches[0].strip().lower()
                
                # Normalisiere auf Standard-Schwierigkeitsgrade
                if any(word in difficulty_text for word in ['leicht', 'einfach', 'easy', 'simple', '1']):
                    return 'Leicht'
                elif any(word in difficulty_text for word in ['schwer', 'difficult', 'hard', 'komplex', 'complex', '3']):
                    return 'Schwer'
                elif any(word in difficulty_text for word in ['mittel', 'medium', 'intermediate', '2']):
                    return 'Mittel'
                else:
                    # Fallback: Versuche direkte Zuordnung
                    if difficulty_text in ['leicht', 'mittel', 'schwer']:
                        return difficulty_text.capitalize()
        
        return None  # Keine explizite Schwierigkeit gefunden

    def _has_inconsistent_difficulty(self, difficulty_raw):
        """Prüft, ob ein Rohwert mehrere Schwierigkeitsgrade gleichzeitig enthält."""
        if not difficulty_raw:
            return False

        value = str(difficulty_raw).lower()
        token_count = 0
        for token in ('leicht', 'mittel', 'schwer'):
            if token in value:
                token_count += 1

        return token_count > 1
    
    def create_document_from_tasks(self, tasks, output_path, lek_theme="", debug_context_report=False):
        """
        Erstellt ein neues Word-Dokument aus einer Liste von Aufgaben
        Verwendet Vorlagen aus dem Vorlagen-Ordner falls verfügbar
        
        Args:
            tasks (list): Liste von Aufgaben-Dictionaries
            output_path (str): Pfad für die neue Datei
            lek_theme (str): LEK-Thema für Vorlagen-Auswahl
            debug_context_report (bool): Gibt Kontext-Report zu Styles/Nummerierung aus
        """
        try:
            # Debug kann explizit oder per Umgebungsvariable aktiviert werden
            env_debug = os.getenv("LEK_DEBUG_EXPORT_CONTEXT", "").strip().lower() in {"1", "true", "yes", "on"}
            debug_context_report = bool(debug_context_report or env_debug)

            # Suche nach passender Vorlage
            template_path = self.template_manager.find_matching_template(lek_theme)
            
            if template_path:
                # Verwende Vorlage als Basis
                doc = self.template_manager.load_template_as_base(template_path)
                print(f"Verwende Vorlage: {os.path.basename(template_path)}")

                # Quellkontext (Styles/Nummerierung) in Ziel-Dokument übernehmen,
                # damit Listen/Formatierungen aus Aufgaben in der Vorlage korrekt bleiben.
                self._prepare_target_document_context(doc, tasks)
                if debug_context_report:
                    self._print_context_report()
                
                # Ersetze Platzhalter wie 'TitelFürThema' durch das LEK-Thema
                self.template_manager.replace_template_placeholders(doc, lek_theme)
                
                # Aufgaben ab Seite 2 einfügen
                self.template_manager.insert_tasks_from_page_2(doc, tasks, lek_theme)
                
            else:
                # Fallback: Erstelle neues Dokument mit programmatischem Deckblatt
                doc = Document()
                
                # Seitenränder einstellen
                sections = doc.sections
                for section in sections:
                    section.top_margin = Inches(1.5 / 2.54)      # 1,5 cm oben
                    section.bottom_margin = Inches(1.5 / 2.54)   # 1,5 cm unten  
                    section.left_margin = Inches(2.5 / 2.54)     # 2,5 cm links
                    section.right_margin = Inches(1.5 / 2.54)    # 1,5 cm rechts
                
                # Standard-Absatzformat anpassen
                styles = doc.styles
                normal_style = styles['Normal']
                normal_font = normal_style.font
                normal_font.name = 'Aptos'
                normal_font.size = Pt(11)
                
                normal_paragraph_format = normal_style.paragraph_format
                normal_paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                normal_paragraph_format.line_spacing = 1.1
                normal_paragraph_format.space_before = Pt(0)
                normal_paragraph_format.space_after = Pt(0)
                
                # Deckblatt erstellen
                self._create_cover_page(doc, lek_theme)
                
                # Neue Seite für den Inhalt
                doc.add_page_break()
                
                # Seitennummerierung ab Seite 2 einrichten (beginnt mit 1)
                self._add_page_numbering_from_page_2(doc)
                
                # Titel der Aufgaben hinzufügen
                title = doc.add_heading('Ihre Lernerfolgskontrolle', 0)
                
                # LEK-Thema als Untertitel hinzufügen, falls vorhanden
                if lek_theme:
                    theme_paragraph = doc.add_paragraph()
                    theme_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    theme_run = theme_paragraph.add_run(lek_theme)
                    theme_run.font.name = 'Aptos'
                    theme_run.font.size = Pt(16)
                    theme_run.font.color.rgb = RGBColor(0, 164, 153)  # Primärfarbe
                    theme_run.bold = True
                    theme_paragraph.paragraph_format.space_after = Pt(12)
                
                # Jede Aufgabe hinzufügen
                for i, task in enumerate(tasks, 1):
                    # Aufgabeninhalt mit vollständiger Struktur - verwende LEK-spezifische Kopier-Logik
                    try:
                        if task.get('all_elements'):
                            # Beste Option: Kopiere alle Elemente mit LEK-angepassten Überschrift-Levels
                            self._copy_elements_for_lek(doc, task['all_elements'])
                        elif task.get('original_paragraphs'):
                            # Sichere Fallback: Kopiere nur Paragraphen mit LEK-Anpassungen
                            self._copy_paragraphs_for_lek(doc, task['original_paragraphs'])
                        else:
                            # Letzter Fallback: Erstelle aus content Array
                            content_list = task.get('content', [])
                            for content_line in content_list:
                                if content_line.strip():
                                    doc.add_paragraph(content_line)
                    except Exception as e:
                        # Notfall-Fallback: Verwende original_paragraphs wenn elements-Kopierung fehlschlägt
                        print(f"Warnung: Element-Kopierung fehlgeschlagen für Aufgabe {i}: {e}")
                        if task.get('original_paragraphs'):
                            self._copy_paragraphs_for_lek(doc, task['original_paragraphs'])
                        else:
                            # Allerletzter Fallback
                            content_list = task.get('content', [])
                            for content_line in content_list:
                                if content_line.strip():
                                    doc.add_paragraph(content_line)
                    
                    # Metadaten als Kommentar (optional)
                    if task.get('difficulty') or task.get('keywords'):
                        meta_info = []
                        if task.get('difficulty'):
                            meta_info.append(f"Schwierigkeit: {task['difficulty']}")
                        if task.get('keywords'):
                            meta_info.append(f"Schlüsselwörter: {', '.join(task['keywords'][:5])}")
                        
                        meta_paragraph = doc.add_paragraph()
                        meta_paragraph.style = normal_style
                        meta_run = meta_paragraph.add_run(' | '.join(meta_info))
                        meta_run.italic = True
                    
                    # Trennlinie zwischen Aufgaben
                    if i < len(tasks):
                        separator_paragraph = doc.add_paragraph("─" * 50)
                        separator_paragraph.style = normal_style
            
            # Dokument speichern
            doc.save(output_path)
            
        except Exception as e:
            raise Exception(f"Fehler beim Erstellen der Word-Datei: {str(e)}")

    def _prepare_target_document_context(self, target_doc, tasks):
        """
        Übernimmt benötigte Styles und Nummerierungsdefinitionen aus dem
        Quell-Dokument der Aufgaben in das Ziel-Dokument.

        Args:
            target_doc: Ziel-Word-Dokument
            tasks: Aufgabenliste mit all_elements
        """
        source_part = self._get_source_part_from_tasks(tasks)
        if source_part is None:
            self._num_id_map = {}
            self._last_context_report = {
                'used_style_ids': [],
                'used_num_ids': [],
                'missing_styles_before': [],
                'missing_styles_after': [],
                'num_id_map': {}
            }
            return

        used_style_ids, used_num_ids = self._collect_used_style_and_num_ids(tasks)
        missing_before = self._get_missing_style_ids(target_doc.part, used_style_ids)
        self._merge_required_styles(source_part, target_doc.part, used_style_ids)
        missing_after = self._get_missing_style_ids(target_doc.part, used_style_ids)
        self._num_id_map = self._merge_required_numbering(source_part, target_doc.part, used_num_ids)

        self._last_context_report = {
            'used_style_ids': sorted(used_style_ids),
            'used_num_ids': sorted(used_num_ids, key=lambda x: int(x) if str(x).isdigit() else 10**9),
            'missing_styles_before': sorted(missing_before),
            'missing_styles_after': sorted(missing_after),
            'num_id_map': dict(self._num_id_map)
        }

    def _get_missing_style_ids(self, target_part, used_style_ids):
        """Ermittelt, welche verwendeten Style-IDs im Ziel noch fehlen."""
        try:
            target_styles = target_part.styles.element
            target_ids = {
                s.get(qn('w:styleId'))
                for s in target_styles.xpath('./w:style')
                if s.get(qn('w:styleId'))
            }
        except Exception:
            return set(used_style_ids)

        return {style_id for style_id in used_style_ids if style_id not in target_ids}

    def _print_context_report(self):
        """Gibt den zuletzt erstellten Kontext-Report aus."""
        report = self._last_context_report or {}
        used_styles = report.get('used_style_ids', [])
        used_nums = report.get('used_num_ids', [])
        missing_before = report.get('missing_styles_before', [])
        missing_after = report.get('missing_styles_after', [])
        num_map = report.get('num_id_map', {})

        print("[Export-Kontext] Styles verwendet      :", len(used_styles))
        print("[Export-Kontext] NumIDs verwendet      :", len(used_nums))
        print("[Export-Kontext] Styles fehlend vorher :", len(missing_before))
        print("[Export-Kontext] Styles fehlend nachher:", len(missing_after))

        if missing_before:
            print("[Export-Kontext] Fehlend vorher (Beispiel):", ', '.join(missing_before[:10]))
        if num_map:
            mappings = [f"{k}->{v}" for k, v in sorted(num_map.items(), key=lambda kv: int(kv[0]) if str(kv[0]).isdigit() else 10**9)]
            print("[Export-Kontext] NumID-Mapping:", ', '.join(mappings[:15]))

    def get_last_context_report(self):
        """Liefert den zuletzt berechneten Export-Kontext-Report zurück."""
        return dict(self._last_context_report) if self._last_context_report else None

    def _get_source_part_from_tasks(self, tasks):
        """
        Ermittelt den Dokument-Part des Quell-Dokuments über die ersten
        verfügbaren Paragraph-/Table-Objekte in den Aufgaben.
        """
        for task in tasks or []:
            for element in task.get('all_elements', []):
                content = element.get('content')
                if hasattr(content, 'part'):
                    return content.part
        return None

    def _collect_used_style_and_num_ids(self, tasks):
        """
        Sammelt verwendete Style-IDs und Num-IDs aus den Aufgaben-Elementen.
        """
        style_ids = set()
        num_ids = set()

        def _xml_of(element):
            content = element.get('content')
            if hasattr(content, '_element'):
                return content._element
            return content

        for task in tasks or []:
            for element in task.get('all_elements', []):
                xml_elem = _xml_of(element)
                if xml_elem is None:
                    continue

                for n in xml_elem.xpath('.//w:pStyle|.//w:rStyle|.//w:tblStyle'):
                    val = n.get(qn('w:val'))
                    if val:
                        style_ids.add(val)

                for n in xml_elem.xpath('.//w:numPr/w:numId'):
                    val = n.get(qn('w:val'))
                    if val:
                        num_ids.add(val)

        return style_ids, num_ids

    def _merge_required_styles(self, source_part, target_part, used_style_ids):
        """
        Fügt fehlende Styles (inkl. Abhängigkeiten) aus Quelle in Ziel ein.
        """
        from copy import deepcopy

        if not used_style_ids:
            return

        try:
            source_styles = source_part.styles.element
            target_styles = target_part.styles.element
        except Exception:
            return

        src_by_id = {
            s.get(qn('w:styleId')): s
            for s in source_styles.xpath('./w:style')
            if s.get(qn('w:styleId'))
        }
        tgt_ids = {
            s.get(qn('w:styleId'))
            for s in target_styles.xpath('./w:style')
            if s.get(qn('w:styleId'))
        }

        pending = list(used_style_ids)
        visited = set()

        while pending:
            style_id = pending.pop()
            if not style_id or style_id in visited:
                continue
            visited.add(style_id)

            src_style = src_by_id.get(style_id)
            if src_style is None:
                continue

            # Abhängigkeiten zuerst aufnehmen
            for dep_xpath in ('./w:basedOn', './w:link', './w:next'):
                dep_nodes = src_style.xpath(dep_xpath)
                for dep in dep_nodes:
                    dep_id = dep.get(qn('w:val'))
                    if dep_id and dep_id not in visited:
                        pending.append(dep_id)

            if style_id not in tgt_ids:
                target_styles.append(deepcopy(src_style))
                tgt_ids.add(style_id)

    def _merge_required_numbering(self, source_part, target_part, used_num_ids):
        """
        Fügt benötigte Nummerierungsdefinitionen aus Quelle in Ziel ein.
        Bei ID-Konflikten werden neue IDs erzeugt und zurückgemappt.

        Returns:
            dict: Mapping alter numId -> neue/übernommene numId
        """
        from copy import deepcopy

        num_id_map = {}
        if not used_num_ids:
            return num_id_map

        try:
            src_numbering = source_part.numbering_part.element
            tgt_numbering = target_part.numbering_part.element
        except Exception:
            return num_id_map

        src_nums = {
            n.get(qn('w:numId')): n
            for n in src_numbering.xpath('./w:num')
            if n.get(qn('w:numId'))
        }
        tgt_nums = {
            n.get(qn('w:numId')): n
            for n in tgt_numbering.xpath('./w:num')
            if n.get(qn('w:numId'))
        }

        src_abs = {
            a.get(qn('w:abstractNumId')): a
            for a in src_numbering.xpath('./w:abstractNum')
            if a.get(qn('w:abstractNumId'))
        }
        tgt_abs = {
            a.get(qn('w:abstractNumId')): a
            for a in tgt_numbering.xpath('./w:abstractNum')
            if a.get(qn('w:abstractNumId'))
        }

        def _next_free_id(existing_ids):
            max_id = max([int(i) for i in existing_ids if str(i).isdigit()] + [0])
            return str(max_id + 1)

        for num_id in sorted(used_num_ids, key=lambda x: int(x) if str(x).isdigit() else 10**9):
            src_num = src_nums.get(num_id)
            if src_num is None:
                continue

            abs_nodes = src_num.xpath('./w:abstractNumId')
            if not abs_nodes:
                continue
            src_abs_id = abs_nodes[0].get(qn('w:val'))

            target_num_id = num_id
            if target_num_id in tgt_nums:
                # ID bereits belegt -> remappen
                target_num_id = _next_free_id(tgt_nums.keys())

            target_abs_id = src_abs_id
            if target_abs_id in tgt_abs:
                # Abstract-ID belegt -> remappen
                target_abs_id = _next_free_id(tgt_abs.keys())

            # AbstractNum sicherstellen
            src_abs_def = src_abs.get(src_abs_id)
            if src_abs_def is not None:
                new_abs = deepcopy(src_abs_def)
                new_abs.set(qn('w:abstractNumId'), target_abs_id)
                if target_abs_id not in tgt_abs:
                    tgt_numbering.append(new_abs)
                    tgt_abs[target_abs_id] = new_abs

            # Num hinzufügen
            new_num = deepcopy(src_num)
            new_num.set(qn('w:numId'), target_num_id)
            new_abs_nodes = new_num.xpath('./w:abstractNumId')
            if new_abs_nodes:
                new_abs_nodes[0].set(qn('w:val'), target_abs_id)

            if target_num_id not in tgt_nums:
                tgt_numbering.append(new_num)
                tgt_nums[target_num_id] = new_num

            num_id_map[num_id] = target_num_id

        return num_id_map

    def _remap_num_ids_in_element(self, xml_element, num_id_map):
        """
        Ersetzt numId-Werte innerhalb eines kopierten XML-Elements anhand Mapping.
        """
        if not num_id_map:
            return

        for node in xml_element.xpath('.//w:numPr/w:numId'):
            old_id = node.get(qn('w:val'))
            if old_id in num_id_map:
                node.set(qn('w:val'), num_id_map[old_id])
    
    def _add_page_numbering_from_page_2(self, doc):
        """
        Fügt Seitennummerierung ab Seite 2 hinzu (beginnend mit 1)
        Format: "Seite/Gesamtseiten" rechtsbündig in der Fußzeile
        
        Args:
            doc: Word-Dokument
        """
        from docx.oxml.shared import qn
        from docx.oxml import OxmlElement
        
        # Erste Sektion (Deckblatt) - keine Seitennummern
        first_section = doc.sections[0]
        first_section.different_first_page_header_footer = True
        
        # Neue Sektion für Inhalt ab Seite 2
        if len(doc.sections) == 1:
            # Falls nur eine Sektion existiert, füge eine neue hinzu
            new_section = doc.add_section()
        else:
            new_section = doc.sections[-1]
        
        # Fußzeile für die neue Sektion konfigurieren
        footer = new_section.footer
        footer_paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        # Seitennummerierung hinzufügen (Feld-Code für PAGE und NUMPAGES)
        run = footer_paragraph.add_run()
        
        # XML für Seitennummerierung erstellen
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        run._r.append(fldChar1)
        
        instrText = OxmlElement('w:instrText')
        instrText.text = 'PAGE \\* MERGEFORMAT'
        run._r.append(instrText)
        
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        run._r.append(fldChar2)
        
        # " / " Trenner
        footer_paragraph.add_run(' / ')
        
        # Gesamtseitenzahl
        run2 = footer_paragraph.add_run()
        
        fldChar3 = OxmlElement('w:fldChar')
        fldChar3.set(qn('w:fldCharType'), 'begin')
        run2._r.append(fldChar3)
        
        instrText2 = OxmlElement('w:instrText')
        instrText2.text = 'NUMPAGES \\* MERGEFORMAT'
        run2._r.append(instrText2)
        
        fldChar4 = OxmlElement('w:fldChar')
        fldChar4.set(qn('w:fldCharType'), 'end')
        run2._r.append(fldChar4)
        
        # Formatierung der Fußzeile
        for run in footer_paragraph.runs:
            run.font.size = Pt(10)
            run.font.name = 'Aptos'
    
    def get_document_info(self, file_path):
        """
        Gibt grundlegende Informationen über ein Word-Dokument zurück
        
        Args:
            file_path (str): Pfad zur Word-Datei
            
        Returns:
            dict: Dokumentinformationen
        """
        try:
            doc = Document(file_path)
            
            paragraph_count = len([p for p in doc.paragraphs if p.text.strip()])
            
            # Geschätzte Aufgabenanzahl basierend auf erkannten Mustern
            task_count = 0
            for paragraph in doc.paragraphs:
                if self._is_task_start(paragraph.text.strip()):
                    task_count += 1
            
            return {
                'paragraph_count': paragraph_count,
                'estimated_task_count': task_count,
                'file_size': f"{len(doc.paragraphs)} Absätze"
            }
            
        except Exception as e:
            raise Exception(f"Fehler beim Analysieren der Datei: {str(e)}")
    
    def _extract_table_text(self, table_element, doc):
        """
        Extrahiert Text aus einer Tabelle für Keyword-Analyse
        
        Args:
            table_element: XML-Element der Tabelle
            doc: Word-Dokument-Objekt
            
        Returns:
            str: Extrahierter Tabellentext
        """
        try:
            # Finde entsprechendes Table-Objekt
            for table in doc.tables:
                if table._element == table_element:
                    text_parts = []
                    for row in table.rows:
                        for cell in row.cells:
                            cell_text = cell.text.strip()
                            if cell_text:
                                text_parts.append(cell_text)
                    return ' '.join(text_parts)
            return ""
        except Exception:
            return ""