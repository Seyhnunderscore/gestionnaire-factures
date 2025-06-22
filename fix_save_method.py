#!/usr/bin/env python3
"""
Script pour corriger l'appel à la méthode save() qui devrait être save_database()
"""

import os
import shutil
from datetime import datetime

def fix_save_method():
    """Corrige l'appel self.database.save() en self.database.save_database()"""
    
    # Chemin vers le fichier main.py
    main_file = "main.py"
    backup_file = f"main.py.backup_save_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
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
        
        # Remplacer self.database.save() par self.database.save_database()
        old_call = "self.database.save()"
        new_call = "self.database.save_database()"
        
        if old_call in content:
            content = content.replace(old_call, new_call)
            print(f"[OK] Remplacement de '{old_call}' par '{new_call}'")
            
            # Compter le nombre de remplacements
            count = content.count(new_call)
            print(f"[INFO] {count} occurrence(s) trouvée(s)")
        else:
            print(f"[INFO] Aucune occurrence de '{old_call}' trouvée")
            return True  # Pas d'erreur si déjà corrigé
        
        # Écrire le contenu modifié
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"[OK] Fichier corrigé avec succès!")
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
    """Vérifie que la correction a été appliquée"""
    main_file = "main.py"
    
    if not os.path.exists(main_file):
        print("Erreur: Impossible de vérifier, le fichier main.py n'existe pas")
        return False
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Vérifier qu'il n'y a plus d'appels à self.database.save()
        if "self.database.save()" in content:
            print("[ATTENTION] Il reste encore des appels à self.database.save()")
            return False
        elif "self.database.save_database()" in content:
            print("[OK] Vérification réussie: utilise maintenant save_database()")
            return True
        else:
            print("[INFO] Aucun appel de sauvegarde trouvé")
            return True
            
    except Exception as e:
        print(f"Erreur lors de la vérification: {str(e)}")
        return False

if __name__ == "__main__":
    print("=== Correction de la méthode de sauvegarde ===")
    print()
    print("Problème identifié:")
    print("- self.database.save() n'existe pas")
    print("- La classe Database a une méthode save_database()")
    print("- Erreur: 'Database' object has no attribute 'save'")
    print()
    print("Solution:")
    print("- Remplacer self.database.save() par self.database.save_database()")
    print()
    
    # Appliquer la correction
    if fix_save_method():
        print()
        print("=== Vérification de la correction ===")
        verify_fix()
        print()
        print("=== Résumé ===")
        print("La méthode de sauvegarde a été corrigée.")
        print("Les modifications de la base de données devraient maintenant")
        print("être sauvegardées correctement sans erreur!")
        print()
        print("Relancez l'application pour tester la correction!")
    else:
        print("Échec de la correction. Vérifiez les messages d'erreur ci-dessus.")
