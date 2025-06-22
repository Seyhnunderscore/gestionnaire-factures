#!/usr/bin/env python3
"""
Script pour augmenter le nombre de lignes parcourues dans la recherche des numéros de facture
"""

import os
import shutil
import re
from datetime import datetime
import logging

def fix_row_range():
    """
    Étend la plage de lignes pour la recherche des numéros de facture
    """
    main_file = "main.py"
    backup_file = f"main.py.backup_rows_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    if not os.path.exists(main_file):
        print("Erreur: main.py non trouvé")
        return False
    
    print(f"Sauvegarde: {backup_file}")
    shutil.copy2(main_file, backup_file)
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 1. Modifier la plage de lignes pour la recherche avec win32com
        old_row_range = 'for row in range(1, 200):'  # Modifié précédemment de 100 à 200
        new_row_range = 'for row in range(1, sheet.UsedRange.Rows.Count + 1):'
        
        content = content.replace(old_row_range, new_row_range)
        
        # 2. Modifier la plage de lignes pour la recherche avec openpyxl
        old_row_range_openpyxl = 'for row in range(1, min(200, sheet.max_row + 1)):'  # Modifié précédemment de 100 à 200
        new_row_range_openpyxl = 'for row in range(1, sheet.max_row + 1):'
        
        content = content.replace(old_row_range_openpyxl, new_row_range_openpyxl)
        
        # 3. Ajouter un log pour indiquer le nombre de lignes analysées
        old_log_sheet = 'logger.info(f"Feuille correspondant à l\'UH {uh} trouvée: {sheet.Name}, {sheet.UsedRange.Columns.Count} colonnes")'
        new_log_sheet = 'logger.info(f"Feuille correspondant à l\'UH {uh} trouvée: {sheet.Name}, {sheet.UsedRange.Columns.Count} colonnes, {sheet.UsedRange.Rows.Count} lignes")'
        
        content = content.replace(old_log_sheet, new_log_sheet)
        
        old_log_sheet_openpyxl = 'logger.info(f"Feuille correspondant à l\'UH {uh} trouvée: {sheet.title}, {sheet.max_column} colonnes")'
        new_log_sheet_openpyxl = 'logger.info(f"Feuille correspondant à l\'UH {uh} trouvée: {sheet.title}, {sheet.max_column} colonnes, {sheet.max_row} lignes")'
        
        content = content.replace(old_log_sheet_openpyxl, new_log_sheet_openpyxl)
        
        # 4. Améliorer la méthode est_facture_correspondante pour être plus souple avec les numéros
        old_method_pattern = r'def est_facture_correspondante\(self, cell_text, facture_num\):.*?return False'
        new_method = '''def est_facture_correspondante(self, cell_text, facture_num):
        """
        Vérifie si une cellule contient le numéro de facture donné,
        avec une approche très souple pour trouver les correspondances
        
        Args:
            cell_text (str): Texte de la cellule à vérifier
            facture_num (str): Numéro de facture à rechercher
        
        Returns:
            bool: True si la cellule contient le numéro de facture, False sinon
        """
        if cell_text is None or facture_num is None:
            return False
            
        # Convertir en texte
        cell_text = str(cell_text).strip()
        facture_num = str(facture_num).strip()
        
        # Log pour diagnostic
        logger.debug(f"Comparaison: '{cell_text}' avec '{facture_num}'")
        
        # Vérification stricte: les textes sont identiques
        if cell_text.lower() == facture_num.lower():
            logger.debug("Correspondance exacte trouvée")
            return True
            
        # Extraire le numéro pur de la facture (après "Facture N°" ou similaire)
        facture_num_pur = facture_num.lower()
        for prefix in ["facture n°", "facture n ", "facture ", "fact ", "fact. ", "n°", "n "]:
            if facture_num_pur.startswith(prefix):
                facture_num_pur = facture_num_pur.replace(prefix, "", 1)
        facture_num_pur = facture_num_pur.strip()
        
        # Nettoyer le texte de la cellule pour comparaison
        cell_text_clean = cell_text.lower()
        for prefix in ["facture n°", "facture n ", "facture ", "fact ", "fact. ", "n°", "n "]:
            if cell_text_clean.startswith(prefix):
                cell_text_clean = cell_text_clean.replace(prefix, "", 1)
        cell_text_clean = cell_text_clean.strip()
        
        # Vérifier si la cellule contient exactement le numéro pur
        if cell_text_clean == facture_num_pur:
            logger.debug(f"Correspondance avec numéro pur: '{facture_num_pur}'")
            return True
            
        # Vérifier si la cellule contient le numéro pur avec "Facture" ou autre préfixe
        if any(cell_text.lower() == f"{prefix}{facture_num_pur}" for prefix in ["facture n°", "facture n ", "facture ", "fact ", "fact. ", "n°", "n "]):
            logger.debug(f"Correspondance avec préfixe et numéro pur")
            return True
            
        # Vérifier si le texte exact de la facture est contenu dans la cellule
        if facture_num.lower() in cell_text.lower():
            logger.debug(f"Facture trouvée comme sous-chaîne: '{facture_num}' dans '{cell_text}'")
            return True
            
        # Vérifier si c'est un nombre qui correspond
        try:
            # Si le numéro pur est un nombre et correspond à la cellule convertie en nombre
            if facture_num_pur.isdigit() and cell_text_clean.replace(".", "").isdigit():
                if int(facture_num_pur) == int(float(cell_text_clean)):
                    logger.debug(f"Correspondance numérique: {facture_num_pur} == {int(float(cell_text_clean))}")
                    return True
        except (ValueError, TypeError):
            pass
            
        # Pas de correspondance trouvée
        return False'''
        
        # Utiliser une expression régulière pour remplacer la méthode
        import re
        content = re.sub(old_method_pattern, new_method, content, flags=re.DOTALL)
        
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("[OK] Plage de lignes étendue pour la recherche des numéros de facture!")
        return True
        
    except Exception as e:
        print(f"Erreur: {e}")
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
        return False

if __name__ == "__main__":
    print("=== Extension de la plage de lignes pour la recherche des numéros de facture ===")
    print()
    print("Problème identifié:")
    print("La feuille Excel contient 669 lignes mais la recherche est limitée à 200 lignes")
    print("- Facture N° 66 se trouve à la cellule P321 (ligne 321)")
    print("- Facture N° 74 se trouve à la cellule P361 (ligne 361)")
    print()
    
    if fix_row_range():
        print()
        print("=== Succès ===")
        print("Les améliorations suivantes ont été apportées:")
        print("1. La recherche parcourt maintenant toutes les lignes utilisées dans la feuille Excel")
        print("2. Les logs indiquent le nombre total de lignes dans chaque feuille")
        print("3. La méthode de comparaison des numéros de facture a été rendue plus souple")
        print()
        print("Relancez l'application pour tester!")
    else:
        print("Échec de la correction.")
