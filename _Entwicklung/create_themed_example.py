"""
Erstellt eine Beispiel-Datei mit korrektem Namensformat für das LEK-Thema
"""

import shutil
import os

def create_themed_example():
    """Kopiert die Beispieldatei und benennt sie im korrekten Format um"""
    
    source = os.path.join('examples', 'beispiel_aufgaben.docx')
    target = os.path.join('examples', 'Aufgaben_Mathematik-Grundlagen.docx')
    
    if os.path.exists(source):
        shutil.copy2(source, target)
        print(f"Themen-Beispieldatei erstellt: {target}")
        print("Diese Datei demonstriert das korrekte Namensformat:")
        print("- LEK-Thema: 'Mathematik-Grundlagen'")
        print("- Export wird vorbelegt mit: LEK_Mathematik-Grundlagen_202411.docx")
    else:
        print(f"Quelldatei nicht gefunden: {source}")

if __name__ == "__main__":
    create_themed_example()