#!/usr/bin/env python3
"""
Script pour corriger le problème de synchronisation entre la base de données
et la validation des factures
"""

import os
import shutil
from datetime import datetime

def fix_synchronization_issue():
    """Corrige le problème de synchronisation des données"""
    
    # Chemin vers le fichier main.py
    main_file = "main.py"
    backup_file = f"main.py.backup_sync_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
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
        
        # Trouver la fonction validate_invoice_row et ajouter la synchronisation
        # Chercher la ligne où on accède à self.database.data
        old_code = """            if hasattr(self, 'database') and self.database.data and db_line is not None:
                # Convertir le dictionnaire en liste pour accéder par index
                db_items = list(self.database.data.items())"""
        
        new_code = """            if hasattr(self, 'database') and self.database.data and db_line is not None:
                # CORRECTION: Forcer la sauvegarde des modifications en attente avant validation
                if hasattr(self, 'pending_changes') and self.pending_changes:
                    self.save_pending_changes()
                    logger.info("Synchronisation forcée des données avant validation")
                
                # Convertir le dictionnaire en liste pour accéder par index
                db_items = list(self.database.data.items())"""
        
        if old_code in content:
            content = content.replace(old_code, new_code)
            print("[OK] Correction de synchronisation ajoutée dans validate_invoice_row")
        else:
            print("[ATTENTION] Code cible non trouvé, tentative alternative...")
            
            # Tentative alternative : chercher une partie plus spécifique
            alt_old = "if hasattr(self, 'database') and self.database.data and db_line is not None:"
            if alt_old in content:
                # Trouver la position et insérer le code de synchronisation
                pos = content.find(alt_old)
                if pos != -1:
                    # Trouver la fin de la ligne
                    end_line = content.find('\n', pos)
                    if end_line != -1:
                        # Insérer le code de synchronisation après cette ligne
                        sync_code = """
                # CORRECTION: Forcer la sauvegarde des modifications en attente avant validation
                if hasattr(self, 'pending_changes') and self.pending_changes:
                    self.save_pending_changes()
                    logger.info("Synchronisation forcée des données avant validation")"""
                        
                        content = content[:end_line] + sync_code + content[end_line:]
                        print("[OK] Correction de synchronisation ajoutée (méthode alternative)")
                    else:
                        print("[ERREUR] Impossible de trouver la fin de ligne")
                        return False
                else:
                    print("[ERREUR] Position non trouvée")
                    return False
            else:
                print("[ERREUR] Code de validation non trouvé")
                return False
        
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
        
        # Vérifier que le code de synchronisation est présent
        if "save_pending_changes()" in content and "Synchronisation forcée" in content:
            print("[OK] Vérification réussie: le code de synchronisation est présent")
            return True
        else:
            print("[ATTENTION] Le code de synchronisation ne semble pas avoir été ajouté")
            return False
            
    except Exception as e:
        print(f"Erreur lors de la vérification: {str(e)}")
        return False

if __name__ == "__main__":
    print("=== Correction du problème de synchronisation ===")
    print()
    print("Problème identifié:")
    print("- Les modifications de la base de données sont sauvegardées avec un délai de 500ms")
    print("- Si l'utilisateur clique sur 'Valider' avant ce délai, les anciennes données sont utilisées")
    print()
    print("Solution:")
    print("- Forcer la sauvegarde des modifications en attente avant chaque validation")
    print()
    
    # Appliquer la correction
    if fix_synchronization_issue():
        print()
        print("=== Vérification de la correction ===")
        verify_fix()
        print()
        print("=== Résumé ===")
        print("Le problème de synchronisation a été corrigé.")
        print("Maintenant, toute modification de la base de données sera")
        print("automatiquement prise en compte lors de la validation des factures.")
        print()
        print("Testez le scénario suivant:")
        print("1. Modifiez une ligne de la base de données")
        print("2. Saisissez immédiatement le numéro de ligne dans les factures")
        print("3. Cliquez sur 'Valider'")
        print("4. Les données modifiées devraient maintenant s'afficher correctement!")
    else:
        print("Échec de la correction. Vérifiez les messages d'erreur ci-dessus.")
