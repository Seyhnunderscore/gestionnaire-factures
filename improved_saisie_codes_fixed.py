#!/usr/bin/env python3
"""
Script pour améliorer la fonction de saisie des codes dans les factures
"""

import os
import shutil
import re
from datetime import datetime
import logging

def clean_facture_number(facture_num):
    """
    Nettoie et normalise un numéro de facture pour faciliter les comparaisons
    
    Args:
        facture_num (str): Numéro de facture à nettoyer
    
    Returns:
        str: Numéro de facture nettoyé
    """
    if facture_num is None:
        return ""
    
    # Convertir en texte et mettre en minuscule
    facture_num = str(facture_num).lower().strip()
    
    # Supprimer les préfixes communs
    prefixes = ["facture n°", "facture n ", "facture ", "fact ", "fact. ", "n°", "n ", "numéro ", "numero "]
    for prefix in prefixes:
        if facture_num.startswith(prefix):
            facture_num = facture_num.replace(prefix, "", 1)
    
    # Supprimer les caractères non alphanumériques tout en conservant les chiffres et lettres
    facture_num = re.sub(r'[^\w\d]', '', facture_num)
    
    return facture_num.strip()

def est_facture_correspondante(cell_text, facture_num):
    """
    Vérifie si une cellule contient un numéro de facture donné,
    en tenant compte de différents formats possibles
    
    Args:
        cell_text (str): Texte de la cellule à vérifier
        facture_num (str): Numéro de facture à rechercher
    
    Returns:
        bool: True si la cellule contient le numéro de facture, False sinon
    """
    if cell_text is None or facture_num is None:
        return False
    
    # Nettoyer les deux numéros pour la comparaison
    clean_cell = clean_facture_number(cell_text)
    clean_facture = clean_facture_number(facture_num)
    
    # Si l'un des numéros est vide après nettoyage, pas de correspondance
    if not clean_cell or not clean_facture:
        return False
    
    # Vérifier si le numéro nettoyé est exactement le même
    if clean_cell == clean_facture:
        return True
    
    # Vérifier si le numéro de facture est contenu dans le texte de la cellule
    # Utile pour les cas où la cellule contient plus d'informations
    if clean_facture in clean_cell:
        return True
    
    # Vérifier si le texte de la cellule est contenu dans le numéro de facture
    # Utile pour les cas où le numéro de facture est plus complet
    if clean_cell in clean_facture:
        return True
    
    return False

def improve_save_invoice_file():
    """
    Améliore la fonction de saisie des codes dans le fichier main.py
    """
    main_file = "main.py"
    backup_file = f"main.py.backup_improved_saisie_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    if not os.path.exists(main_file):
        print("Erreur: main.py non trouvé")
        return False
    
    print(f"Sauvegarde: {backup_file}")
    shutil.copy2(main_file, backup_file)
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Remplacer la méthode est_facture_correspondante par notre version améliorée
        old_method_pattern = r'def est_facture_correspondante\(self, cell_text, facture_num\):.*?return False'
        new_method = '''def est_facture_correspondante(self, cell_text, facture_num):
        """
        Vérifie si une cellule contient un numéro de facture donné,
        en tenant compte de différents formats possibles
        
        Args:
            cell_text (str): Texte de la cellule à vérifier
            facture_num (str): Numéro de facture à rechercher
        
        Returns:
            bool: True si la cellule contient le numéro de facture, False sinon
        """
        if cell_text is None or facture_num is None:
            return False
        
        # Nettoyer les deux numéros pour la comparaison
        def clean_facture_number(num):
            if num is None:
                return ""
            
            # Convertir en texte et mettre en minuscule
            num = str(num).lower().strip()
            
            # Supprimer les préfixes communs
            prefixes = ["facture n°", "facture n ", "facture ", "fact ", "fact. ", "n°", "n ", "numéro ", "numero "]
            for prefix in prefixes:
                if num.startswith(prefix):
                    num = num.replace(prefix, "", 1)
            
            # Supprimer les caractères non alphanumériques tout en conservant les chiffres et lettres
            num = re.sub(r'[^\\\\w\\\\d]', '', num)
            
            return num.strip()
        
        # Nettoyer les deux numéros pour la comparaison
        clean_cell = clean_facture_number(cell_text)
        clean_facture = clean_facture_number(facture_num)
        
        # Si l'un des numéros est vide après nettoyage, pas de correspondance
        if not clean_cell or not clean_facture:
            return False
        
        # Vérifier si le numéro nettoyé est exactement le même
        if clean_cell == clean_facture:
            return True
        
        # Vérifier si le numéro de facture est contenu dans le texte de la cellule
        # Utile pour les cas où la cellule contient plus d'informations
        if clean_facture in clean_cell:
            return True
        
        # Vérifier si le texte de la cellule est contenu dans le numéro de facture
        # Utile pour les cas où le numéro de facture est plus complet
        if clean_cell in clean_facture:
            return True
        
        return False'''
        
        # Utiliser une expression régulière pour remplacer la méthode
        import re
        content = re.sub(old_method_pattern, new_method, content, flags=re.DOTALL)
        
        # Modifier la méthode save_invoice_file pour utiliser est_facture_correspondante
        # au lieu de la comparaison stricte
        old_comparison = r'if cell_num_pur == facture_num_pur:'
        new_comparison = 'if self.est_facture_correspondante(cell_text, facture_num):'
        
        # Remplacer les deux occurrences (win32com et openpyxl)
        content = content.replace(old_comparison, new_comparison)
        
        # Augmenter les limites de recherche pour couvrir plus de cellules
        content = content.replace('for row in range(1, 100):', 'for row in range(1, 200):')  # Augmenter à 200 lignes
        content = content.replace('for col in range(1, 20):', 'for col in range(1, 30):')  # Augmenter à 30 colonnes
        
        content = content.replace('for row in range(1, min(100, sheet.max_row + 1)):', 'for row in range(1, min(200, sheet.max_row + 1)):')
        content = content.replace('for col in range(1, min(20, sheet.max_column + 1)):', 'for col in range(1, min(30, sheet.max_column + 1)):')
        
        # Ajouter un message de log plus détaillé pour les factures non trouvées
        old_warning = 'logger.warning(f"Numéro de facture {facture_num} non trouvé dans la feuille correspondant à l\'UH {uh}")'
        new_warning = 'logger.warning(f"Numéro de facture {facture_num} (nettoyé: {facture_num_pur}) non trouvé dans la feuille correspondant à l\'UH {uh}")'
        
        content = content.replace(old_warning, new_warning)
        
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("[OK] Fonction saisie des codes améliorée!")
        return True
        
    except Exception as e:
        print(f"Erreur: {e}")
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
        return False

if __name__ == "__main__":
    print("=== Amélioration de la fonction de saisie des codes ===")
    print()
    print("Problèmes identifiés:")
    print("1. Comparaison trop stricte des numéros de facture")
    print("2. Limites de recherche trop restrictives (100 lignes, 20 colonnes)")
    print("3. Traitement des numéros de facture insuffisant")
    print()
    
    if improve_save_invoice_file():
        print()
        print("=== Succès ===")
        print("La fonction de saisie des codes a été améliorée pour:")
        print("- Mieux détecter les correspondances entre numéros de facture")
        print("- Rechercher dans plus de lignes et colonnes (200 lignes, 30 colonnes)")
        print("- Fournir des logs plus détaillés en cas d'échec")
        print()
        print("Relancez l'application pour tester!")
    else:
        print("Échec de l'amélioration.")
