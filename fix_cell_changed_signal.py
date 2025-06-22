#!/usr/bin/env python3
"""
Script pour corriger le problème avec on_db_cell_changed qui ne reçoit pas les bons paramètres
"""

import os
import shutil
from datetime import datetime

def fix_cell_changed_signal():
    """Corrige la fonction on_db_cell_changed pour qu'elle fonctionne avec itemChanged"""
    
    # Chemin vers le fichier main.py
    main_file = "main.py"
    backup_file = f"main.py.backup_signal_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
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
        
        # Remplacer la fonction on_db_cell_changed pour qu'elle fonctionne avec itemChanged
        old_function = """    def on_db_cell_changed(self, row, column):
        \"\"\"Appelé lorsqu'une cellule du tableau est modifiée\"\"\"
        if self._updating_table or not hasattr(self, 'db_table'):
            return
            
        try:
            # Récupérer le nom du client (colonne 1)
            name_item = self.db_table.item(row, 1)
            if not name_item or not name_item.text().strip():
                return
                
            name = name_item.text().strip()
            
            # Mettre à jour les modifications en attente
            self.pending_changes[name] = {
                'client_code': self.db_table.item(row, 2).text() if self.db_table.item(row, 2) else "",
                'chorus_code': self.db_table.item(row, 3).text() if self.db_table.item(row, 3) else "",
                'address': self.db_table.item(row, 4).text() if self.db_table.item(row, 4) else ""
            }
            
            # Démarrer/redémarrer le minuteur de sauvegarde (500ms)
            if hasattr(self, 'save_timer'):
                self.save_timer.start(500)
                
            logger.debug(f"Modification enregistrée pour {name}")
            
        except Exception as e:
            logger.error(f"Erreur lors de l'enregistrement de la modification: {e}")"""
        
        new_function = """    def on_db_cell_changed(self, item):
        \"\"\"Appelé lorsqu'une cellule du tableau est modifiée\"\"\"
        if self._updating_table or not hasattr(self, 'db_table') or not item:
            return
            
        try:
            # Récupérer la ligne et colonne de l'item modifié
            row = item.row()
            column = item.column()
            
            # Récupérer le nom du client (colonne 1)
            name_item = self.db_table.item(row, 1)
            if not name_item or not name_item.text().strip():
                return
                
            name = name_item.text().strip()
            
            # Mettre à jour les modifications en attente
            self.pending_changes[name] = {
                'client_code': self.db_table.item(row, 2).text() if self.db_table.item(row, 2) else "",
                'chorus_code': self.db_table.item(row, 3).text() if self.db_table.item(row, 3) else "",
                'address': self.db_table.item(row, 4).text() if self.db_table.item(row, 4) else ""
            }
            
            # Démarrer/redémarrer le minuteur de sauvegarde (500ms)
            if hasattr(self, 'save_timer'):
                self.save_timer.start(500)
                
            logger.debug(f"Modification enregistrée pour {name} (ligne {row}, colonne {column})")
            
        except Exception as e:
            logger.error(f"Erreur lors de l'enregistrement de la modification: {e}")"""
        
        if old_function in content:
            content = content.replace(old_function, new_function)
            print("[OK] Fonction on_db_cell_changed corrigée pour itemChanged")
        else:
            print("[ATTENTION] Fonction exacte non trouvée, tentative alternative...")
            
            # Tentative alternative : chercher juste la signature de la fonction
            old_signature = "def on_db_cell_changed(self, row, column):"
            new_signature = "def on_db_cell_changed(self, item):"
            
            if old_signature in content:
                content = content.replace(old_signature, new_signature)
                
                # Remplacer aussi les références à row et column dans la fonction
                # Chercher le début de la fonction et remplacer le contenu
                func_start = content.find(new_signature)
                if func_start != -1:
                    # Trouver la fin de la fonction (prochaine fonction ou fin de classe)
                    func_end = content.find("\n    def ", func_start + 1)
                    if func_end == -1:
                        func_end = content.find("\nclass ", func_start + 1)
                    if func_end == -1:
                        func_end = len(content)
                    
                    # Remplacer le contenu de la fonction
                    func_content = content[func_start:func_end]
                    
                    # Ajouter les lignes pour récupérer row et column depuis item
                    if "row = item.row()" not in func_content:
                        insert_pos = func_content.find("try:")
                        if insert_pos != -1:
                            insert_pos += 4  # après "try:"
                            new_lines = """
            # Récupérer la ligne et colonne de l'item modifié
            row = item.row()
            column = item.column()
            """
                            func_content = func_content[:insert_pos] + new_lines + func_content[insert_pos:]
                            content = content[:func_start] + func_content + content[func_end:]
                
                print("[OK] Signature de fonction corrigée (méthode alternative)")
            else:
                print("[ERREUR] Impossible de trouver la fonction on_db_cell_changed")
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
        
        # Vérifier que la nouvelle signature est présente
        if "def on_db_cell_changed(self, item):" in content:
            print("[OK] Vérification réussie: la fonction utilise maintenant le bon paramètre")
            return True
        else:
            print("[ATTENTION] La signature de fonction ne semble pas avoir été corrigée")
            return False
            
    except Exception as e:
        print(f"Erreur lors de la vérification: {str(e)}")
        return False

if __name__ == "__main__":
    print("=== Correction du signal itemChanged ===")
    print()
    print("Problème identifié:")
    print("- on_db_cell_changed(self, row, column) attend 2 paramètres")
    print("- itemChanged ne passe qu'un seul paramètre (l'item modifié)")
    print("- Cela cause l'erreur: 'missing 1 required positional argument: column'")
    print()
    print("Solution:")
    print("- Modifier la fonction pour accepter un paramètre 'item'")
    print("- Récupérer row et column depuis item.row() et item.column()")
    print()
    
    # Appliquer la correction
    if fix_cell_changed_signal():
        print()
        print("=== Vérification de la correction ===")
        verify_fix()
        print()
        print("=== Résumé ===")
        print("Le signal itemChanged devrait maintenant fonctionner correctement.")
        print("Les modifications de la base de données seront détectées et sauvegardées.")
        print()
        print("Relancez l'application pour tester la correction!")
    else:
        print("Échec de la correction. Vérifiez les messages d'erreur ci-dessus.")
