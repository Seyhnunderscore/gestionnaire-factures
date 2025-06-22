#!/usr/bin/env python3
"""
Script pour corriger le filtre "sans couleur" et le texte "sans concordance"
"""

import os
import shutil
from datetime import datetime

def fix_sans_couleur():
    """Corrige le filtre sans couleur et le texte"""
    
    # Chemin vers le fichier main.py
    main_file = "main.py"
    backup_file = f"main.py.backup_fix_sans_couleur_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    # Vérifier que le fichier existe
    if not os.path.exists(main_file):
        print(f"Erreur: Le fichier {main_file} n'existe pas")
        return False
    
    # Créer une sauvegarde
    print(f"Création d'une sauvegarde: {backup_file}")
    shutil.copy2(main_file, backup_file)
    
    try:
        # Lire le contenu du fichier
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 1. Corriger le texte dans le menu déroulant
        old_text = 'self.color_filter_combo.addItem("Pas de correspondance (Sans couleur)")'
        new_text = 'self.color_filter_combo.addItem("Pas de concordance (Sans couleur)")'
        
        if old_text in content:
            content = content.replace(old_text, new_text)
            print("[OK] Texte 'correspondance' → 'concordance' corrigé")
        
        # 2. Corriger la logique de détection dans filter_invoices_by_color
        old_condition = 'elif selected_filter == "Pas de correspondance (Sans couleur)":'
        new_condition = 'elif selected_filter == "Pas de concordance (Sans couleur)":'
        
        if old_condition in content:
            content = content.replace(old_condition, new_condition)
            print("[OK] Condition de filtre corrigée")
        
        # 3. Améliorer la logique de détection des cellules sans couleur
        old_sans_couleur_logic = """                    elif selected_filter == "Pas de concordance (Sans couleur)":
                        # Pas de couleur de fond (blanc ou transparent)
                        color_name = background_color.name().lower()
                        color_match = (color_name in ["#ffffff", "#f0f0f0"] or 
                                     background_color.alpha() == 0 or
                                     (background_color.red() > 240 and 
                                      background_color.green() > 240 and 
                                      background_color.blue() > 240))"""
        
        new_sans_couleur_logic = """                    elif selected_filter == "Pas de concordance (Sans couleur)":
                        # Pas de couleur de fond - cellules par défaut
                        color_name = background_color.name().lower()
                        # Vérifier si c'est une couleur par défaut (blanc, transparent, ou très clair)
                        is_default_color = (
                            color_name in ["#ffffff", "#f0f0f0", "#000000"] or
                            background_color.alpha() == 0 or
                            # Couleur très claire (proche du blanc)
                            (background_color.red() > 240 and 
                             background_color.green() > 240 and 
                             background_color.blue() > 240) or
                            # Vérifier si ce n'est PAS une des couleurs spécifiques
                            (color_name != "#0078d7" and  # Pas bleu
                             not (background_color.red() < 150 and background_color.green() > 150 and background_color.blue() < 150) and  # Pas vert
                             not (background_color.red() > 200 and background_color.green() > 100 and background_color.green() < 200 and background_color.blue() < 100))  # Pas orange
                        )
                        color_match = is_default_color"""
        
        if old_sans_couleur_logic in content:
            content = content.replace(old_sans_couleur_logic, new_sans_couleur_logic)
            print("[OK] Logique de détection 'sans couleur' améliorée")
        else:
            print("[ATTENTION] Logique 'sans couleur' non trouvée pour correction")
        
        # 4. Aussi corriger dans la condition else (cas où il n'y a pas d'item)
        old_else_condition = 'color_match = (selected_filter == "Pas de correspondance (Sans couleur)")'
        new_else_condition = 'color_match = (selected_filter == "Pas de concordance (Sans couleur)")'
        
        if old_else_condition in content:
            content = content.replace(old_else_condition, new_else_condition)
            print("[OK] Condition else corrigée")
        
        # Écrire le contenu modifié
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"[OK] Corrections appliquées avec succès!")
        print(f"[OK] Sauvegarde créée: {backup_file}")
        return True
        
    except Exception as e:
        print(f"Erreur lors de la correction: {str(e)}")
        # Restaurer la sauvegarde en cas d'erreur
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
            print(f"[OK] Fichier restauré depuis la sauvegarde")
        return False

if __name__ == "__main__":
    print("=== Correction du filtre 'Sans couleur' ===")
    print()
    print("Corrections à apporter:")
    print("- Texte: 'correspondance' → 'concordance'")
    print("- Amélioration de la détection des cellules sans couleur")
    print("- Logique inverse: détecter ce qui N'EST PAS coloré")
    print()
    
    # Appliquer la correction
    if fix_sans_couleur():
        print()
        print("=== Résumé ===")
        print(" Texte corrigé: 'Pas de concordance (Sans couleur)'")
        print(" Logique améliorée pour détecter les cellules sans couleur")
        print(" Détection par exclusion des couleurs spécifiques")
        print()
        print("Relancez l'application pour tester les corrections!")
    else:
        print("Échec de la correction. Vérifiez les messages d'erreur ci-dessus.")
