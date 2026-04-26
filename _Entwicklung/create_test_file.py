"""
Erstellt eine Test-Aufgabendatei für Auftragssteuerung und -koordination
"""

import shutil
import os

def create_test_file():
    """Erstellt eine Test-Aufgabendatei mit passendem Namen"""
    
    source = os.path.join('examples', 'beispiel_aufgaben.docx')
    target = os.path.join('examples', 'Aufgaben_Auftragssteuerung-Koordination.docx')
    
    if os.path.exists(source):
        shutil.copy2(source, target)
        print(f"Test-Aufgabendatei erstellt: {target}")
        print("Diese Datei wird automatisch mit der passenden Vorlage verknüpft.")
    else:
        print(f"Quelldatei nicht gefunden: {source}")

if __name__ == "__main__":
    create_test_file()