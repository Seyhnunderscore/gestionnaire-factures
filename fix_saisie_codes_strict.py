#!/usr/bin/env python3
"""
Script pour corriger la fonction de saisie des codes avec une comparaison plus stricte
"""

import os
import shutil
import re
from datetime import datetime
import logging

def fix_saisie_codes_strict():
    """
    Corrige la fonction de saisie des codes pour utiliser une comparaison plus stricte
    """
    main_file = "main.py"
    backup_file = f"main.py.backup_strict_saisie_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    if not os.path.exists(main_file):
        print("Erreur: main.py non trouvé")
        return False
    
    print(f"Sauvegarde: {backup_file}")
    shutil.copy2(main_file, backup_file)
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Remplacer la méthode est_facture_correspondante par notre version stricte
        old_method_pattern = r'def est_facture_correspondante\(self, cell_text, facture_num\):.*?return False'
        new_method = '''def est_facture_correspondante(self, cell_text, facture_num):
        """
        Vérifie si une cellule contient exactement le numéro de facture donné,
        en tenant compte uniquement des variations de format standard
        
        Args:
            cell_text (str): Texte de la cellule à vérifier
            facture_num (str): Numéro de facture à rechercher
        
        Returns:
            bool: True si la cellule contient exactement le numéro de facture, False sinon
        """
        if cell_text is None or facture_num is None:
            return False
        
        # Convertir en texte
        cell_text = str(cell_text).strip()
        facture_num = str(facture_num).strip()
        
        # Vérification stricte: les textes doivent être identiques
        if cell_text.lower() == facture_num.lower():
            return True
        
        # Formats standards pour "Facture N° X"
        cell_formats = [
            cell_text.lower(),
            cell_text.lower().replace("facture n°", "").strip(),
            cell_text.lower().replace("facture n ", "").strip(),
            cell_text.lower().replace("facture ", "").strip()
        ]
        
        facture_formats = [
            facture_num.lower(),
            facture_num.lower().replace("facture n°", "").strip(),
            facture_num.lower().replace("facture n ", "").strip(),
            facture_num.lower().replace("facture ", "").strip()
        ]
        
        # Vérifier si l'un des formats correspond exactement
        for cell_format in cell_formats:
            for facture_format in facture_formats:
                if cell_format == facture_format and cell_format.strip() != "":
                    return True
        
        # Vérifier si c'est un nombre entier qui correspond exactement
        try:
            # Si les deux sont des nombres et sont égaux
            if cell_text.replace(".", "").isdigit() and facture_num.replace(".", "").isdigit():
                cell_num = int(float(cell_text))
                facture_num_int = int(float(facture_num.replace("facture n°", "").replace("facture n ", "").replace("facture ", "")))
                if cell_num == facture_num_int:
                    return True
        except (ValueError, TypeError):
            pass
        
        # Pas de correspondance trouvée
        return False'''
        
        # Utiliser une expression régulière pour remplacer la méthode
        import re
        content = re.sub(old_method_pattern, new_method, content, flags=re.DOTALL)
        
        # Ajouter un log plus détaillé pour les correspondances trouvées
        old_log = 'logger.info(f"Numéro de facture {facture_num} trouvé dans la cellule ({row}, {col}): \'{cell_text}\'")'
        new_log = 'logger.info(f"Numéro de facture {facture_num} trouvé dans la cellule ({row}, {col}): \'{cell_text}\' - VÉRIFICATION: formats identiques")'
        
        content = content.replace(old_log, new_log)
        
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("[OK] Fonction saisie des codes corrigée avec comparaison stricte!")
        return True
        
    except Exception as e:
        print(f"Erreur: {e}")
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
        return False

if __name__ == "__main__":
    print("=== Correction stricte de la fonction de saisie des codes ===")
    print()
    print("Problème identifié:")
    print("La méthode actuelle trouve des correspondances partielles incorrectes")
    print("Exemple: 'Facture N° 74' est trouvé dans '25741.0'")
    print()
    
    if fix_saisie_codes_strict():
        print()
        print("=== Succès ===")
        print("La fonction de saisie des codes a été corrigée pour:")
        print("- Utiliser une comparaison stricte des numéros de facture")
        print("- Ne détecter que les correspondances exactes")
        print("- Gérer correctement les formats standards (Facture N° X, etc.)")
        print()
        print("Relancez l'application pour tester!")
    else:
        print("Échec de la correction.")
