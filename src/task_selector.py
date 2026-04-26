"""
TaskSelector - Klasse für die Auswahl und Filterung von Aufgaben
"""

import random
import re

class TaskSelector:
    """Klasse für die Auswahl von Aufgaben basierend auf verschiedenen Kriterien"""
    
    def filter_tasks(self, tasks, criteria):
        """
        Filtert Aufgaben basierend auf den angegebenen Kriterien
        
        Args:
            tasks (list): Liste aller verfügbaren Aufgaben
            criteria (dict): Filterkriterien
                - keywords (list): Liste von Suchbegriffen
                - difficulty (str): Gewünschter Schwierigkeitsgrad
                - max_count (int): Maximale Anzahl von Aufgaben
                - min_length (int): Minimale Textlänge
                - max_length (int): Maximale Textlänge
                
        Returns:
            list: Gefilterte Liste von Aufgaben
        """
        filtered = tasks.copy()
        
        # Nach Schlüsselwörtern filtern
        if criteria.get('keywords'):
            filtered = self._filter_by_keywords(filtered, criteria['keywords'])
        
        # Nach Schwierigkeitsgrad filtern
        if criteria.get('difficulty'):
            filtered = self._filter_by_difficulty(filtered, criteria['difficulty'])
        
        # Nach Textlänge filtern
        if criteria.get('min_length'):
            filtered = self._filter_by_min_length(filtered, criteria['min_length'])
        
        if criteria.get('max_length'):
            filtered = self._filter_by_max_length(filtered, criteria['max_length'])
        
        # Anzahl begrenzen
        max_count = criteria.get('max_count', len(filtered))
        if len(filtered) > max_count:
            # Zufällige Auswahl für Vielfalt
            filtered = random.sample(filtered, max_count)
        
        return filtered
    
    def _filter_by_keywords(self, tasks, keywords):
        """
        Filtert Aufgaben, die mindestens eines der Schlüsselwörter enthalten
        
        Args:
            tasks (list): Liste von Aufgaben
            keywords (list): Liste von Suchbegriffen
            
        Returns:
            list: Gefilterte Aufgaben
        """
        if not keywords:
            return tasks
        
        filtered_tasks = []
        
        for task in tasks:
            task_text = ' '.join(task.get('content', [])).lower()
            task_keywords = [kw.lower() for kw in task.get('keywords', [])]
            task_title = task.get('title', '').lower()
            
            # Prüfen ob eines der Schlüsselwörter im Text, Titel oder Keywords vorkommt
            match_found = False
            for keyword in keywords:
                keyword_lower = keyword.lower().strip()
                if (keyword_lower in task_text or 
                    keyword_lower in task_title or
                    any(keyword_lower in tk for tk in task_keywords)):
                    match_found = True
                    break
            
            if match_found:
                # Relevanz-Score für bessere Sortierung
                score = self._calculate_keyword_relevance(task, keywords)
                task['relevance_score'] = score
                filtered_tasks.append(task)
        
        # Nach Relevanz sortieren
        filtered_tasks.sort(key=lambda x: x.get('relevance_score', 0), reverse=True)
        
        return filtered_tasks
    
    def _calculate_keyword_relevance(self, task, keywords):
        """
        Berechnet einen Relevanz-Score basierend auf Keyword-Übereinstimmungen
        
        Args:
            task (dict): Aufgaben-Dictionary
            keywords (list): Liste von Suchbegriffen
            
        Returns:
            int: Relevanz-Score (höher = relevanter)
        """
        score = 0
        task_text = ' '.join(task.get('content', [])).lower()
        task_title = task.get('title', '').lower()
        task_keywords = [kw.lower() for kw in task.get('keywords', [])]
        
        for keyword in keywords:
            keyword_lower = keyword.lower().strip()
            
            # Punkte für Titel-Treffer (höchste Priorität)
            if keyword_lower in task_title:
                score += 10
            
            # Punkte für exakte Keyword-Treffer
            if keyword_lower in task_keywords:
                score += 5
            
            # Punkte für Text-Treffer
            score += task_text.count(keyword_lower) * 2
        
        return score
    
    def _filter_by_difficulty(self, tasks, difficulty):
        """
        Filtert Aufgaben nach Schwierigkeitsgrad
        
        Args:
            tasks (list): Liste von Aufgaben
            difficulty (str): Gewünschter Schwierigkeitsgrad
            
        Returns:
            list: Gefilterte Aufgaben
        """
        return [task for task in tasks if task.get('difficulty') == difficulty]
    
    def _filter_by_min_length(self, tasks, min_length):
        """
        Filtert Aufgaben nach minimaler Textlänge
        
        Args:
            tasks (list): Liste von Aufgaben
            min_length (int): Minimale Anzahl Zeichen
            
        Returns:
            list: Gefilterte Aufgaben
        """
        filtered = []
        for task in tasks:
            total_length = sum(len(content) for content in task.get('content', []))
            if total_length >= min_length:
                filtered.append(task)
        return filtered
    
    def _filter_by_max_length(self, tasks, max_length):
        """
        Filtert Aufgaben nach maximaler Textlänge
        
        Args:
            tasks (list): Liste von Aufgaben
            max_length (int): Maximale Anzahl Zeichen
            
        Returns:
            list: Gefilterte Aufgaben
        """
        filtered = []
        for task in tasks:
            total_length = sum(len(content) for content in task.get('content', []))
            if total_length <= max_length:
                filtered.append(task)
        return filtered
    
    def suggest_keywords(self, tasks):
        """
        Schlägt relevante Schlüsselwörter basierend auf den verfügbaren Aufgaben vor
        
        Args:
            tasks (list): Liste aller Aufgaben
            
        Returns:
            list: Liste der häufigsten Schlüsselwörter
        """
        keyword_frequency = {}
        
        for task in tasks:
            for keyword in task.get('keywords', []):
                keyword_lower = keyword.lower()
                keyword_frequency[keyword_lower] = keyword_frequency.get(keyword_lower, 0) + 1
        
        # Sortiert nach Häufigkeit, mindestens 2 Vorkommen
        suggested = [
            keyword for keyword, freq in keyword_frequency.items() 
            if freq >= 2
        ]
        
        return sorted(suggested, key=lambda x: keyword_frequency[x], reverse=True)[:20]
    
    def get_difficulty_distribution(self, tasks):
        """
        Gibt die Verteilung der Schwierigkeitsgrade zurück
        
        Args:
            tasks (list): Liste von Aufgaben
            
        Returns:
            dict: Dictionary mit Schwierigkeitsgrad-Verteilung
        """
        distribution = {}
        for task in tasks:
            difficulty = task.get('difficulty', 'Unbekannt')
            distribution[difficulty] = distribution.get(difficulty, 0) + 1
        
        return distribution
    
    def create_balanced_selection(self, tasks, total_count):
        """
        Erstellt eine ausgewogene Auswahl von Aufgaben verschiedener Schwierigkeitsgrade
        
        Args:
            tasks (list): Liste aller verfügbaren Aufgaben
            total_count (int): Gewünschte Gesamtanzahl
            
        Returns:
            list: Ausgewogene Auswahl von Aufgaben
        """
        # Aufgaben nach Schwierigkeitsgrad gruppieren
        difficulty_groups = {
            'Leicht': [t for t in tasks if t.get('difficulty') == 'Leicht'],
            'Mittel': [t for t in tasks if t.get('difficulty') == 'Mittel'],
            'Schwer': [t for t in tasks if t.get('difficulty') == 'Schwer']
        }
        
        # Gleichmäßige Verteilung anstreben
        selected = []
        per_difficulty = total_count // 3
        remainder = total_count % 3
        
        for difficulty, group in difficulty_groups.items():
            count = per_difficulty
            if remainder > 0:
                count += 1
                remainder -= 1
            
            if group:
                selected.extend(random.sample(group, min(count, len(group))))
        
        # Falls noch Plätze frei sind, mit zufälligen Aufgaben auffüllen
        if len(selected) < total_count:
            remaining_tasks = [t for t in tasks if t not in selected]
            if remaining_tasks:
                additional = random.sample(remaining_tasks, 
                                         min(total_count - len(selected), len(remaining_tasks)))
                selected.extend(additional)
        
        return selected[:total_count]