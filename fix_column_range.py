#!/usr/bin/env python3
"""
Script pour corriger la plage de colonnes dans la recherche des numéros de facture
"""

import os
import shutil
import re
from datetime import datetime
import logging

def fix_column_range():
    """
    Étend la plage de colonnes pour la recherche des numéros de facture
    """
    main_file = "main.py"
    backup_file = f"main.py.backup_columns_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    if not os.path.exists(main_file):
        print("Erreur: main.py non trouvé")
        return False
    
    print(f"Sauvegarde: {backup_file}")
    shutil.copy2(main_file, backup_file)
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 1. Modifier la plage de colonnes pour la recherche avec win32com
        old_col_range = 'for col in range(1, 30):'  # Modifié précédemment de 20 à 30
        new_col_range = 'for col in range(1, sheet.UsedRange.Columns.Count + 1):'
        
        content = content.replace(old_col_range, new_col_range)
        
        # 2. Modifier la plage de colonnes pour la recherche avec openpyxl
        old_col_range_openpyxl = 'for col in range(1, min(30, sheet.max_column + 1)):'  # Modifié précédemment de 20 à 30
        new_col_range_openpyxl = 'for col in range(1, sheet.max_column + 1):'
        
        content = content.replace(old_col_range_openpyxl, new_col_range_openpyxl)
        
        # 3. Ajouter un log pour indiquer le nombre de colonnes analysées
        old_log_sheet = 'logger.info(f"Feuille correspondant à l\'UH {uh} trouvée: {sheet.Name}")'
        new_log_sheet = 'logger.info(f"Feuille correspondant à l\'UH {uh} trouvée: {sheet.Name}, {sheet.UsedRange.Columns.Count} colonnes")'
        
        content = content.replace(old_log_sheet, new_log_sheet)
        
        old_log_sheet_openpyxl = 'logger.info(f"Feuille correspondant à l\'UH {uh} trouvée: {sheet.title}")'
        new_log_sheet_openpyxl = 'logger.info(f"Feuille correspondant à l\'UH {uh} trouvée: {sheet.title}, {sheet.max_column} colonnes")'
        
        content = content.replace(old_log_sheet_openpyxl, new_log_sheet_openpyxl)
        
        # 4. Ajouter un log pour indiquer la position exacte où le code est inséré
        old_log_found = 'logger.info(f"Numéro de facture {facture_num} trouvé dans la cellule ({row}, {col}): \'{cell_text}\' - VÉRIFICATION: formats identiques")'
        new_log_found = 'logger.info(f"Numéro de facture {facture_num} trouvé dans la cellule ({row}, {col}, colonne {chr(64 + col) if col <= 26 else chr(64 + col//26) + chr(64 + col%26)}): \'{cell_text}\' - VÉRIFICATION: formats identiques")'
        
        content = content.replace(old_log_found, new_log_found)
        
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("[OK] Plage de colonnes étendue pour la recherche des numéros de facture!")
        return True
        
    except Exception as e:
        print(f"Erreur: {e}")
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
        return False

if __name__ == "__main__":
    print("=== Extension de la plage de colonnes pour la recherche des numéros de facture ===")
    print()
    print("Problème identifié:")
    print("Les numéros de facture sont dans la colonne P (16) mais la recherche est limitée")
    print("- Facture N° 66 se trouve à la cellule P321")
    print("- Facture N° 74 se trouve à la cellule P361")
    print()
    
    if fix_column_range():
        print()
        print("=== Succès ===")
        print("Les améliorations suivantes ont été apportées:")
        print("1. La recherche parcourt maintenant toutes les colonnes utilisées dans la feuille Excel")
        print("2. Les logs indiquent le nombre total de colonnes dans chaque feuille")
        print("3. Les logs indiquent la référence exacte de la cellule (ex: P321) où le numéro de facture est trouvé")
        print()
        print("Relancez l'application pour tester!")
    else:
        print("Échec de la correction.")
