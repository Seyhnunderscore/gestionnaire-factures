#!/usr/bin/env python3
"""
Script pour corriger le problème d'importation Excel dans main.py
Le problème : la fonction import_database appelle self.update_database_view() qui n'existe pas
La solution : remplacer par self.load_database_into_table()
"""

import os
import shutil
from datetime import datetime

def fix_import_excel_function():
    """Corrige la fonction import_database pour l'importation Excel"""
    
    # Chemin vers le fichier main.py
    main_file = "main.py"
    backup_file = f"main.py.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
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
        
        # Remplacements à effectuer
        replacements = [
            # Ligne 1747 : remplacer update_database_view par load_database_into_table
            ("self.update_database_view()", "self.load_database_into_table()"),
            
            # Autres corrections potentielles dans la même fonction
            ("self.update_database_view", "self.load_database_into_table"),
            ("self.refresh_database_table", "self.load_database_into_table"),
            ("self.populate_database_table", "self.load_database_into_table")
        ]
        
        # Appliquer les remplacements
        modifications_made = 0
        for old_text, new_text in replacements:
            if old_text in content:
                content = content.replace(old_text, new_text)
                modifications_made += 1
                print(f"[OK] Remplace: {old_text} -> {new_text}")
        
        if modifications_made == 0:
            print("Aucune modification nécessaire trouvée")
            return True
        
        # Écrire le contenu modifié
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"[OK] Fichier corrigé avec succès! {modifications_made} modification(s) appliquée(s)")
        print(f"[OK] Sauvegarde créée: {backup_file}")
        return True
        
    except Exception as e:
        print(f"Erreur lors de la correction: {str(e)}")
        # Restaurer la sauvegarde en cas d'erreur
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
            print(f"[OK] Fichier restauré depuis la sauvegarde")
        return False

def verify_fix():
    """Vérifie que la correction a été appliquée correctement"""
    main_file = "main.py"
    
    if not os.path.exists(main_file):
        print("Erreur: Impossible de vérifier, le fichier main.py n'existe pas")
        return False
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Vérifier que les problèmes ont été corrigés
        problems = []
        if "self.update_database_view()" in content:
            problems.append("self.update_database_view() toujours présent")
        
        if problems:
            print("[ATTENTION] Problèmes détectés après correction:")
            for problem in problems:
                print(f"  - {problem}")
            return False
        else:
            print("[OK] Vérification réussie: toutes les corrections ont été appliquées")
            return True
            
    except Exception as e:
        print(f"Erreur lors de la vérification: {str(e)}")
        return False

if __name__ == "__main__":
    print("=== Correction du problème d'importation Excel ===")
    print()
    
    # Appliquer la correction
    if fix_import_excel_function():
        print()
        print("=== Vérification de la correction ===")
        verify_fix()
        print()
        print("=== Résumé ===")
        print("Le problème d'importation Excel a été corrigé.")
        print("La fonction import_database utilise maintenant load_database_into_table()")
        print("pour mettre à jour l'affichage après l'importation d'un fichier Excel.")
        print()
        print("Vous pouvez maintenant tester l'importation de fichiers Excel!")
    else:
        print("Échec de la correction. Vérifiez les messages d'erreur ci-dessus.")
