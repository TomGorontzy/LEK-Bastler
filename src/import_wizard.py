"""
Import Wizard - Zustandsmodell für geführten Aufgabenimport.

Sprint 1 (Schritt 1):
- Datenmodelle für ImportTask und ImportSession
- Session-Statistiken und Freigabe-Helfer
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any
import re


Confidence = str  # "high" | "medium" | "low"


def _as_string_list(value: Any) -> list[str]:
    """Normalisiert beliebige Eingaben robust auf eine Liste von Strings."""
    if value is None:
        return []

    if isinstance(value, str):
        item = value.strip()
        return [item] if item else []

    if isinstance(value, (list, tuple, set)):
        normalized: list[str] = []
        for item in value:
            text = str(item or "").strip()
            if text:
                normalized.append(text)
        return normalized

    text = str(value).strip()
    return [text] if text else []


def is_blocking_warning(text: str) -> bool:
    """Klassifiziert Warnungen, die in der Regel den Export blockieren sollten."""
    warning_text = str(text or "").strip().lower()
    if not warning_text:
        return False

    blocking_markers = (
        "pflichtfeld fehlt",
        "kategorie fehlt",
        "inkonsistenter schwierigkeitsgrad",
        "keinen verwertbaren inhalt",
    )
    return any(marker in warning_text for marker in blocking_markers)


def _blocking_warning_key(text: str) -> str | None:
    """Ordnet eine blockierende Warnung einem stabilen Schlüssel zu."""
    warning_text = str(text or "").strip().lower()
    if not warning_text:
        return None

    if "pflichtfeld fehlt" in warning_text:
        if "titel" in warning_text:
            return "missing_title"
        return "missing_required"
    if "kategorie fehlt" in warning_text:
        return "missing_category"
    if "inkonsistenter schwierigkeitsgrad" in warning_text:
        return "inconsistent_difficulty"
    if "keinen verwertbaren inhalt" in warning_text:
        return "missing_content"

    return None


@dataclass
class ImportTask:
    """Repräsentiert eine importierte Aufgabe inkl. Diagnosemetadaten."""

    id: int
    number_display: str
    category: str
    title: str
    intro: list[str] = field(default_factory=list)
    content_elements: list[Any] = field(default_factory=list)
    difficulty: str = "Mittel"
    keywords: list[str] = field(default_factory=list)
    confidence: Confidence = "medium"
    warnings: list[str] = field(default_factory=list)
    accepted: bool = True
    raw_task: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def from_raw_task(
        cls,
        task: dict[str, Any],
        default_confidence: Confidence = "medium",
        fallback_id: int = 0,
    ) -> "ImportTask":
        task_number_raw = task.get("number", fallback_id)
        task_id = cls._to_internal_id(task_number_raw, fallback_id=fallback_id)
        number_display = str(
            task.get("number_display")
            or task.get("number_label")
            or task_number_raw
            or task_id
        ).strip()
        if not number_display:
            number_display = str(task_id)

        category = task.get("category", "Ohne Kategorie")
        title = task.get("title", "Ohne Titel")
        difficulty = task.get("difficulty", "Mittel")
        keywords = _as_string_list(task.get("keywords", []))
        elements = list(task.get("all_elements", []) or [])
        intro = _as_string_list(task.get("intro", []))
        warnings = _as_string_list(task.get("warnings", []))
        confidence = task.get("confidence", default_confidence)

        return cls(
            id=task_id,
            number_display=number_display,
            category=category,
            title=title,
            intro=intro,
            content_elements=elements,
            difficulty=difficulty,
            keywords=keywords,
            confidence=confidence,
            warnings=warnings,
            raw_task=task,
        )

    @staticmethod
    def _to_internal_id(value: Any, fallback_id: int = 0) -> int:
        """Ermittelt eine stabile numerische interne ID (auch bei 1.1/1.2-Nummern)."""
        try:
            return int(value)
        except (TypeError, ValueError):
            pass

        text = str(value or "").strip()
        if text:
            match = re.match(r"^(\d+)", text)
            if match:
                try:
                    return int(match.group(1))
                except ValueError:
                    pass

        try:
            return int(fallback_id)
        except (TypeError, ValueError):
            return 0


@dataclass
class ImportSession:
    """Session-Zustand für den Wizard-Flow."""

    source_file: str
    source_filename: str
    lek_theme: str
    tasks: list[ImportTask] = field(default_factory=list)
    global_warnings: list[str] = field(default_factory=list)
    approved_task_ids: set[int] = field(default_factory=set)

    @classmethod
    def from_raw_tasks(
        cls,
        source_file: str,
        source_filename: str,
        lek_theme: str,
        raw_tasks: list[dict[str, Any]],
    ) -> "ImportSession":
        tasks = [ImportTask.from_raw_task(t, fallback_id=i) for i, t in enumerate(raw_tasks, 1)]
        approved_ids = {t.id for t in tasks if t.accepted and t.id > 0}

        global_warnings: list[str] = []
        if len(tasks) == 0:
            global_warnings.append("Keine Aufgaben erkannt.")

        return cls(
            source_file=source_file,
            source_filename=source_filename,
            lek_theme=lek_theme,
            tasks=tasks,
            global_warnings=global_warnings,
            approved_task_ids=approved_ids,
        )

    def get_stats(self) -> dict[str, int]:
        total = len(self.tasks)
        approved = len(self.approved_task_ids)
        high = sum(1 for t in self.tasks if t.confidence == "high")
        medium = sum(1 for t in self.tasks if t.confidence == "medium")
        low = sum(1 for t in self.tasks if t.confidence == "low")
        warnings = sum(1 for t in self.tasks if t.warnings)
        blocking = sum(
            1
            for t in self.tasks
            if any(is_blocking_warning(w) for w in (t.warnings or []))
        )

        return {
            "total": total,
            "approved": approved,
            "high": high,
            "medium": medium,
            "low": low,
            "warnings": warnings,
            "blocking": blocking,
        }

    def approve_all(self) -> None:
        self.approved_task_ids = {t.id for t in self.tasks if t.id > 0}
        for task in self.tasks:
            task.accepted = task.id in self.approved_task_ids

    def clear_approvals(self) -> None:
        self.approved_task_ids.clear()
        for task in self.tasks:
            task.accepted = False

    def set_task_approval(self, task_id: int, accepted: bool) -> None:
        found = False
        for task in self.tasks:
            if task.id == task_id:
                task.accepted = accepted
                found = True
                break

        if not found:
            return

        if accepted:
            self.approved_task_ids.add(task_id)
        else:
            self.approved_task_ids.discard(task_id)

    def get_approved_raw_tasks(self) -> list[dict[str, Any]]:
        approved_ids = self.approved_task_ids
        return [t.raw_task for t in self.tasks if t.id in approved_ids]

    def get_blocking_summary(self) -> dict[str, int]:
        """Liefert eine kategoriale Zusammenfassung blockierender Warnungen."""
        counts = {
            "missing_title": 0,
            "missing_required": 0,
            "missing_category": 0,
            "inconsistent_difficulty": 0,
            "missing_content": 0,
        }

        for task in self.tasks:
            task_keys = set()
            for warning in task.warnings or []:
                key = _blocking_warning_key(warning)
                if key:
                    task_keys.add(key)

            for key in task_keys:
                if key in counts:
                    counts[key] += 1

        counts["total"] = sum(counts.values())
        return counts
