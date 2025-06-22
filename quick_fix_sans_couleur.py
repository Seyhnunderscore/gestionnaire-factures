#!/usr/bin/env python3
"""
Script simple pour corriger le texte et la logique sans couleur
"""

import os
import shutil
from datetime import datetime

def quick_fix():
    main_file = "main.py"
    backup_file = f"main.py.backup_quick_fix_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    if not os.path.exists(main_file):
        print("Erreur: main.py non trouve")
        return False
    
    print(f"Sauvegarde: {backup_file}")
    shutil.copy2(main_file, backup_file)
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Corrections simples
        content = content.replace('"Pas de correspondance (Sans couleur)"', '"Pas de concordance (Sans couleur)"')
        
        # Améliorer la logique sans couleur - remplacer toute la section
        old_logic = """                    elif selected_filter == "Pas de concordance (Sans couleur)":
                        # Pas de couleur de fond (blanc ou transparent)
                        color_name = background_color.name().lower()
                        color_match = (color_name in ["#ffffff", "#f0f0f0"] or 
                                     background_color.alpha() == 0 or
                                     (background_color.red() > 240 and 
                                      background_color.green() > 240 and 
                                      background_color.blue() > 240))"""
        
        new_logic = """                    elif selected_filter == "Pas de concordance (Sans couleur)":
                        # Cellules sans couleur spécifique (par défaut)
                        color_name = background_color.name().lower()
                        # Une cellule est "sans couleur" si elle n'a PAS les couleurs spécifiques
                        is_blue = color_name == "#0078d7"
                        is_green = (background_color.red() < 150 and background_color.green() > 150 and background_color.blue() < 150)
                        is_orange = (background_color.red() > 200 and background_color.green() > 100 and background_color.green() < 200 and background_color.blue() < 100)
                        
                        # Sans couleur = PAS bleu ET PAS vert ET PAS orange
                        color_match = not (is_blue or is_green or is_orange)"""
        
        if old_logic in content:
            content = content.replace(old_logic, new_logic)
            print("[OK] Logique sans couleur corrigee")
        else:
            print("[ATTENTION] Ancienne logique non trouvee")
        
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("[OK] Corrections appliquees!")
        return True
        
    except Exception as e:
        print(f"Erreur: {e}")
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
        return False

if __name__ == "__main__":
    print("=== Correction rapide ===")
    if quick_fix():
        print("Succes! Relancez l'application.")
    else:
        print("Echec de la correction.")
