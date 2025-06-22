#!/usr/bin/env python3
"""
Script pour ajouter l'import manquant QGridLayout dans main.py
"""

import os
import shutil
from datetime import datetime

def fix_qgridlayout_import():
    """Ajoute QGridLayout à la liste des imports PyQt5"""
    
    # Chemin vers le fichier main.py
    main_file = "main.py"
    backup_file = f"main.py.backup_qgrid_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    # Vérifier que le fichier existe
    if not os.path.exists(main_file):
        print(f"Erreur: Le fichier {main_file} n'existe pas")
        return False
    
    # Créer une sauvegarde
    print(f"Creation d'une sauvegarde: {backup_file}")
    shutil.copy2(main_file, backup_file)
    
    try:
        # Lire le contenu du fichier
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Chercher la ligne avec QFormLayout et ajouter QGridLayout
        old_line = "    QSpinBox, QStyle, QStyleFactory, QFormLayout, QSplitter, QDialog,"
        new_line = "    QSpinBox, QStyle, QStyleFactory, QFormLayout, QGridLayout, QSplitter, QDialog,"
        
        if old_line in content:
            content = content.replace(old_line, new_line)
            print("[OK] QGridLayout ajoute aux imports PyQt5")
            
            # Écrire le contenu modifié
            with open(main_file, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print(f"[OK] Fichier corrige avec succes!")
            print(f"[OK] Sauvegarde creee: {backup_file}")
            return True
        else:
            print("Ligne d'import non trouvee, tentative alternative...")
            
            # Tentative alternative : chercher une autre ligne d'import
            alt_old = "from PyQt5.QtWidgets import ("
            if alt_old in content:
                # Trouver la fin de l'import et ajouter QGridLayout
                import_start = content.find(alt_old)
                import_end = content.find(")", import_start)
                
                if import_end != -1:
                    # Vérifier si QGridLayout est déjà présent
                    import_section = content[import_start:import_end]
                    if "QGridLayout" not in import_section:
                        # Ajouter QGridLayout avant la parenthèse fermante
                        before_closing = content[:import_end]
                        after_closing = content[import_end:]
                        
                        # Ajouter QGridLayout avec une virgule
                        if not before_closing.endswith(",\n    "):
                            before_closing += ",\n    "
                        else:
                            before_closing += ""
                        
                        content = before_closing + "QGridLayout" + after_closing
                        
                        # Écrire le contenu modifié
                        with open(main_file, 'w', encoding='utf-8') as f:
                            f.write(content)
                        
                        print("[OK] QGridLayout ajoute aux imports PyQt5 (methode alternative)")
                        print(f"[OK] Fichier corrige avec succes!")
                        return True
                    else:
                        print("[INFO] QGridLayout est deja present dans les imports")
                        return True
            
            print("Impossible de trouver la section d'import PyQt5")
            return False
        
    except Exception as e:
        print(f"Erreur lors de la correction: {str(e)}")
        # Restaurer la sauvegarde en cas d'erreur
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
            print(f"[OK] Fichier restaure depuis la sauvegarde")
        return False

def verify_fix():
    """Vérifie que QGridLayout est bien importé"""
    main_file = "main.py"
    
    if not os.path.exists(main_file):
        print("Erreur: Impossible de verifier, le fichier main.py n'existe pas")
        return False
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Vérifier que QGridLayout est présent dans les imports
        if "QGridLayout" in content and "from PyQt5.QtWidgets import" in content:
            print("[OK] Verification reussie: QGridLayout est maintenant importe")
            return True
        else:
            print("[ATTENTION] QGridLayout ne semble pas etre importe correctement")
            return False
            
    except Exception as e:
        print(f"Erreur lors de la verification: {str(e)}")
        return False

if __name__ == "__main__":
    print("=== Correction de l'import QGridLayout ===")
    print()
    
    # Appliquer la correction
    if fix_qgridlayout_import():
        print()
        print("=== Verification de la correction ===")
        verify_fix()
        print()
        print("=== Resume ===")
        print("L'import QGridLayout a ete ajoute aux imports PyQt5.")
        print("L'importation Excel devrait maintenant fonctionner sans erreur!")
    else:
        print("Echec de la correction. Verifiez les messages d'erreur ci-dessus.")
