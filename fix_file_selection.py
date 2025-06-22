#!/usr/bin/env python3
"""
Script pour améliorer la fonction de saisie des codes en ajoutant une sélection de fichier
en cas de problème avec le fichier Excel
"""

import os
import shutil
import re
from datetime import datetime
import logging

def fix_file_selection():
    """
    Améliore la fonction save_invoice_file pour permettre à l'utilisateur de sélectionner
    un fichier Excel en cas de problème avec le fichier actuel
    """
    main_file = "main.py"
    backup_file = f"main.py.backup_file_selection_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    if not os.path.exists(main_file):
        print(f"Erreur: {main_file} non trouvé")
        return False
    
    print(f"Sauvegarde: {backup_file}")
    shutil.copy2(main_file, backup_file)
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Rechercher la partie du code qui vérifie si un fichier Excel est chargé
        old_code = """# Vérifier si un fichier Excel est chargé
        if not hasattr(self, 'current_excel_file') or not self.current_excel_file:
            # Si aucun fichier n'est chargé, utiliser le dernier fichier ouvert
            if hasattr(self, 'current_invoice_path') and self.current_invoice_path and os.path.exists(self.current_invoice_path):
                self.current_excel_file = self.current_invoice_path
                logger.info(f"Utilisation du dernier fichier ouvert: {self.current_excel_file}")
            else:
                QMessageBox.warning(self, "Attention", "Aucun fichier de facturation n'est chargé.")
                return False"""
        
        # Nouveau code avec la possibilité de sélectionner un fichier
        new_code = """# Vérifier si un fichier Excel est chargé
        if not hasattr(self, 'current_excel_file') or not self.current_excel_file or not os.path.exists(self.current_excel_file):
            # Si aucun fichier n'est chargé ou si le fichier n'existe plus, utiliser le dernier fichier ouvert
            if hasattr(self, 'current_invoice_path') and self.current_invoice_path and os.path.exists(self.current_invoice_path):
                self.current_excel_file = self.current_invoice_path
                logger.info(f"Utilisation du dernier fichier ouvert: {self.current_excel_file}")
            else:
                # Proposer à l'utilisateur de sélectionner un fichier Excel
                reply = QMessageBox.question(
                    self,
                    "Fichier non trouvé",
                    "Le fichier Excel n'a pas été trouvé ou n'est pas chargé. Voulez-vous sélectionner un fichier Excel?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                
                if reply == QMessageBox.Yes:
                    # Ouvrir une boîte de dialogue pour sélectionner un fichier Excel
                    file_dialog = QFileDialog()
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
                        return False
                else:
                    QMessageBox.warning(self, "Attention", "Aucun fichier de facturation n'est chargé.")
                    return False"""
        
        # Remplacer l'ancien code par le nouveau
        content = content.replace(old_code, new_code)
        
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
            else:
                # Si aucun import de PyQt5.QtWidgets n'est trouvé, ajouter une nouvelle ligne d'import
                content = "from PyQt5.QtWidgets import QFileDialog\n" + content
        
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("[OK] Fonction de sélection de fichier ajoutée!")
        return True
        
    except Exception as e:
        print(f"Erreur: {e}")
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
        return False

if __name__ == "__main__":
    print("=== Ajout de la sélection de fichier pour la saisie des codes ===")
    print()
    print("Problème identifié:")
    print("- Si le fichier Excel change de nom ou d'emplacement, le bouton 'Saisie des codes' ne fonctionne plus")
    print()
    
    if fix_file_selection():
        print()
        print("=== Succès ===")
        print("La fonction de saisie des codes a été améliorée pour:")
        print("- Permettre à l'utilisateur de sélectionner un fichier Excel en cas de problème")
        print("- Vérifier l'existence du fichier avant de l'utiliser")
        print()
        print("Relancez l'application pour tester!")
    else:
        print("Échec de l'amélioration.")
