#!/usr/bin/env python3
"""
Script pour améliorer la fonction de saisie des codes en ajoutant une sélection de fichier
directe en cas de problème avec le fichier Excel
"""

import os
import shutil
import re
from datetime import datetime
import logging

def fix_file_selection_direct():
    """
    Améliore la fonction save_invoice_file pour ouvrir directement l'explorateur de fichiers
    en cas de problème avec le fichier Excel
    """
    main_file = "main.py"
    backup_file = f"main.py.backup_file_selection_direct_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    if not os.path.exists(main_file):
        print(f"Erreur: {main_file} non trouvé")
        return False
    
    print(f"Sauvegarde: {backup_file}")
    shutil.copy2(main_file, backup_file)
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Rechercher toutes les occurrences de ce message d'erreur
        old_message = 'QMessageBox.warning(self, "Attention", "Aucun fichier de facturation n\'est chargé.")'
        
        # Nouveau code qui ouvre directement l'explorateur de fichiers
        new_code = '''# Ouvrir directement une boîte de dialogue pour sélectionner un fichier Excel
                file_dialog = QFileDialog()
                file_dialog.setWindowTitle("Sélectionner un fichier Excel de facturation")
                file_dialog.setFileMode(QFileDialog.ExistingFile)
                file_dialog.setNameFilter("Fichiers Excel (*.xlsx *.xls)")
                file_dialog.setViewMode(QFileDialog.Detail)
                
                if file_dialog.exec_():
                    selected_files = file_dialog.selectedFiles()
                    if selected_files:
                        self.current_excel_file = selected_files[0]
                        logger.info(f"Fichier Excel sélectionné: {self.current_excel_file}")
                    else:
                        QMessageBox.warning(self, "Attention", "Aucun fichier n'a été sélectionné.")
                        return False
                else:
                    QMessageBox.warning(self, "Attention", "Aucun fichier n'a été sélectionné.")
                    return False'''
        
        # Remplacer toutes les occurrences du message d'erreur par le nouveau code
        if old_message in content:
            content = content.replace(old_message, new_code)
            print(f"[OK] Message d'erreur remplacé par l'ouverture directe de l'explorateur de fichiers")
        else:
            print("[ATTENTION] Message d'erreur non trouvé")
            
            # Essayer une autre approche en cherchant un pattern plus général
            pattern = r'QMessageBox\.warning\(self,\s*"Attention",\s*"Aucun fichier[^"]*"\)'
            matches = re.findall(pattern, content)
            
            if matches:
                for match in matches:
                    content = content.replace(match, new_code)
                print(f"[OK] {len(matches)} messages d'erreur remplacés par l'ouverture directe de l'explorateur de fichiers")
            else:
                print("[ERREUR] Aucun message d'erreur correspondant trouvé")
                return False
        
        # Ajouter l'import de QFileDialog s'il n'est pas déjà présent
        if "from PyQt5.QtWidgets import QFileDialog" not in content:
            # Trouver la ligne d'import de PyQt5.QtWidgets
            import_pattern = r"from PyQt5.QtWidgets import (.*)"
            import_match = re.search(import_pattern, content)
            
            if import_match:
                old_import = import_match.group(0)
                # Vérifier si QFileDialog est déjà dans l'import
                if "QFileDialog" not in old_import:
                    # Ajouter QFileDialog à l'import existant
                    new_import = old_import.replace(")", ", QFileDialog)")
                    content = content.replace(old_import, new_import)
                    print("[OK] Import de QFileDialog ajouté")
            else:
                # Si aucun import de PyQt5.QtWidgets n'est trouvé, ajouter une nouvelle ligne d'import
                content = "from PyQt5.QtWidgets import QFileDialog\n" + content
                print("[OK] Nouvelle ligne d'import de QFileDialog ajoutée")
        
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("[OK] Fonction de sélection directe de fichier ajoutée!")
        return True
        
    except Exception as e:
        print(f"Erreur: {e}")
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
        return False

if __name__ == "__main__":
    print("=== Ajout de la sélection directe de fichier pour la saisie des codes ===")
    print()
    print("Problème identifié:")
    print("- L'application affiche toujours 'Aucun fichier de facturation n'est chargé' au lieu d'ouvrir l'explorateur")
    print()
    
    if fix_file_selection_direct():
        print()
        print("=== Succès ===")
        print("La fonction de saisie des codes a été améliorée pour:")
        print("- Ouvrir DIRECTEMENT l'explorateur de fichiers en cas de problème avec le fichier Excel")
        print("- Permettre à l'utilisateur de sélectionner un fichier Excel sans confirmation préalable")
        print()
        print("Relancez l'application pour tester!")
    else:
        print("Échec de l'amélioration.")
