import os
import sys
import tempfile
import unittest
from copy import deepcopy

from docx import Document


ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
SRC = os.path.join(ROOT, 'src')
if SRC not in sys.path:
    sys.path.insert(0, SRC)

from word_processor import WordProcessor
from import_wizard import ImportSession


class RegressionCoreTests(unittest.TestCase):
    def setUp(self):
        self.wp = WordProcessor()
        self.base_rules = deepcopy(self.wp.rules)

    def tearDown(self):
        self.wp.rules = deepcopy(self.base_rules)

    def _make_table_doc(self, path, rows):
        doc = Document()
        table = doc.add_table(rows=0, cols=2)
        for key, value in rows:
            row = table.add_row()
            row.cells[0].text = key
            row.cells[1].text = value
        doc.save(path)

    def _extract_first_task(self, path):
        tasks = self.wp.extract_tasks(path)
        self.assertTrue(tasks, 'Es wurde keine Aufgabe erkannt.')
        return tasks[0]

    def test_table_with_intro_hint_points_preview_order(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, 'table_full.docx')
            self._make_table_doc(path, [
                ('ID', 'A-001'),
                ('Titel', 'Komplettaufgabe'),
                ('Intro/Einleitung (optional)', 'Kontextblock vorhanden'),
                ('Aufgabenstellung (Pflicht)', 'Bitte bearbeiten'),
                ('Lösungsmöglichkeit/Hinweis (optional)', 'Achten Sie auf X'),
                ('Punkte', '10'),
                ('Kategorie', 'Demo'),
                ('Schwierigkeitsgrad', 'mittel'),
            ])

            task = self._extract_first_task(path)
            lines = self.wp.build_task_flow_preview_lines(task)

            joined = '\n'.join(lines)
            expected_order = [
                'Titel:',
                'Intro/Einleitung:',
                'Aufgabenstellung:',
                'Hinweis:',
                'Punkte:',
            ]
            last_pos = -1
            for marker in expected_order:
                pos = joined.find(marker)
                self.assertGreater(pos, last_pos, f"Abschnittsreihenfolge fehlerhaft bei '{marker}'")
                last_pos = pos

    def test_table_without_optional_blocks_delta_check(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, 'table_minimal.docx')
            self._make_table_doc(path, [
                ('ID', 'A-002'),
                ('Titel', 'Minimalaufgabe'),
                ('Aufgabenstellung (Pflicht)', 'Nur Pflichtinhalt'),
                ('Kategorie', 'Demo'),
                ('Schwierigkeitsgrad', 'mittel'),
            ])

            task = self._extract_first_task(path)
            delta = self.wp.analyze_task_flow_preview_delta(task)

            self.assertTrue(delta.get('is_structured'))
            missing = set(delta.get('missing_optional_sections') or [])
            self.assertTrue({'Intro/Einleitung', 'Hinweis', 'Punkte'}.issubset(missing))

    def test_missing_category_warning_and_block_rule(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, 'table_missing_category.docx')
            self._make_table_doc(path, [
                ('ID', 'A-003'),
                ('Titel', 'Ohne Kategorie'),
                ('Aufgabenstellung (Pflicht)', 'Inhalt'),
                ('Kategorie', ''),
                ('Schwierigkeitsgrad', 'mittel'),
            ])

            task = self._extract_first_task(path)
            diagnostics = self.wp._build_task_diagnostic(task)

            warnings = diagnostics.get('warnings') or []
            self.assertTrue(any('Kategorie fehlt' in w for w in warnings))
            self.assertTrue(bool(self.wp.get_import_rule('category_rules.block_export_on_missing', True)))

    def test_missing_title_fallback_and_required_warning(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, 'table_missing_title.docx')
            self._make_table_doc(path, [
                ('ID', 'A-004'),
                ('Aufgabenstellung (Pflicht)', 'Eine ausführliche Aufgabenstellung für Titelableitung'),
                ('Kategorie', 'Demo'),
                ('Schwierigkeitsgrad', 'mittel'),
            ])

            self.wp.rules = deepcopy(self.base_rules)
            self.wp.rules.setdefault('template_rules', {})['required_fields'] = ['id', 'titel', 'aufgabenstellungpflicht']

            task = self._extract_first_task(path)
            self.assertNotEqual((task.get('title') or '').strip(), '')
            self.assertTrue(any('Pflichtfeld fehlt: Titel' in w for w in (task.get('pre_warnings') or [])))

    def test_inconsistent_difficulty_warning_and_block(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, 'table_inconsistent_diff.docx')
            self._make_table_doc(path, [
                ('ID', 'A-005'),
                ('Titel', 'Uneinheitliche Schwierigkeit'),
                ('Aufgabenstellung (Pflicht)', 'Inhalt'),
                ('Kategorie', 'Demo'),
                ('Schwierigkeitsgrad', 'leicht und schwer'),
            ])

            task = self._extract_first_task(path)
            self.assertTrue(any('Inkonsistenter Schwierigkeitsgrad' in w for w in (task.get('pre_warnings') or [])))
            self.assertTrue(bool(self.wp.get_import_rule('difficulty_rules.block_export_on_inconsistent', True)))

    def test_partial_approval_exports_exact_approved(self):
        raw_tasks = [
            {'number': 1, 'number_display': '1', 'category': 'A', 'title': 'T1', 'content': ['c1'], 'difficulty': 'Mittel', 'keywords': []},
            {'number': 2, 'number_display': '2', 'category': 'A', 'title': 'T2', 'content': ['c2'], 'difficulty': 'Mittel', 'keywords': []},
            {'number': 3, 'number_display': '3', 'category': 'A', 'title': 'T3', 'content': ['c3'], 'difficulty': 'Mittel', 'keywords': []},
        ]

        session = ImportSession.from_raw_tasks(
            source_file='dummy.docx',
            source_filename='dummy.docx',
            lek_theme='Demo',
            raw_tasks=raw_tasks,
        )

        session.clear_approvals()
        session.set_task_approval(1, True)
        session.set_task_approval(3, True)

        approved = session.get_approved_raw_tasks()
        approved_numbers = [int(t.get('number', 0)) for t in approved]
        self.assertEqual(approved_numbers, [1, 3])


if __name__ == '__main__':
    unittest.main()