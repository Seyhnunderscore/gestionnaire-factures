#!/usr/bin/env python3
"""
Script pour corriger le filtre par couleur qui ne fonctionne pas
"""

import os
import shutil
from datetime import datetime

def fix_color_filter():
    """Corrige la logique du filtre par couleur"""
    
    # Chemin vers le fichier main.py
    main_file = "main.py"
    backup_file = f"main.py.backup_fix_color_filter_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
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
        
        # Remplacer la méthode filter_invoices complète
        old_filter_method = """    def filter_invoices(self):
        \"\"\"Filtre les factures en fonction de la recherche\"\"\"
        search_text = self.invoice_search_input.text().lower()
        
        for row in range(self.invoice_table.rowCount()):
            hide_row = True
            for col in range(self.invoice_table.columnCount()):
                item = self.invoice_table.item(row, col)
                if item and search_text in item.text().lower():
                    hide_row = False
                    break
            
            self.invoice_table.setRowHidden(row, hide_row)
        
        # Appliquer aussi le filtre par couleur
        self.filter_invoices_by_color()"""
        
        new_filter_method = """    def filter_invoices(self):
        \"\"\"Filtre les factures en fonction de la recherche\"\"\"
        search_text = self.invoice_search_edit.text().lower()
        
        for row in range(self.invoice_table.rowCount()):
            hide_row = True
            for col in range(self.invoice_table.columnCount()):
                item = self.invoice_table.item(row, col)
                if item and search_text in item.text().lower():
                    hide_row = False
                    break
            
            self.invoice_table.setRowHidden(row, hide_row)
        
        # Appliquer aussi le filtre par couleur
        self.filter_invoices_by_color()"""
        
        if old_filter_method in content:
            content = content.replace(old_filter_method, new_filter_method)
            print("[OK] Correction du nom de champ de recherche")
        
        # Remplacer la méthode filter_invoices_by_color complète
        old_color_filter = """    def filter_invoices_by_color(self):
        \"\"\"Filtre les factures en fonction de la couleur sélectionnée\"\"\"
        if not hasattr(self, 'color_filter_combo'):
            return
            
        selected_filter = self.color_filter_combo.currentText()
        
        for row in range(self.invoice_table.rowCount()):
            # Si déjà caché par le filtre texte, ne pas modifier
            if self.invoice_table.isRowHidden(row):
                continue
                
            # Obtenir la couleur de la première cellule de la ligne
            item = self.invoice_table.item(row, 0)
            if not item:
                continue
                
            background_color = item.background().color()
            hide_row = False
            
            if selected_filter == \"Toutes les couleurs\":
                hide_row = False
            elif selected_filter == \"Validé (Bleu)\":
                # Couleur bleue (#0078d7)
                hide_row = background_color.name().lower() != \"#0078d7\"
            elif selected_filter == \"Correspondance exacte (Vert)\":
                # Couleur verte (success color)
                hide_row = not (background_color.red() < 100 and background_color.green() > 150 and background_color.blue() < 100)
            elif selected_filter == \"Correspondance partielle (Orange)\":
                # Couleur orange (warning color)
                hide_row = not (background_color.red() > 200 and background_color.green() > 100 and background_color.blue() < 100)
            elif selected_filter == \"Pas de correspondance (Sans couleur)\":
                # Pas de couleur de fond ou couleur par défaut
                hide_row = background_color.name().lower() not in [\"#ffffff\", \"#f0f0f0\", \"#transparent\"]
            
            self.invoice_table.setRowHidden(row, hide_row)"""
        
        new_color_filter = """    def filter_invoices_by_color(self):
        \"\"\"Filtre les factures en fonction de la couleur sélectionnée\"\"\"
        if not hasattr(self, 'color_filter_combo'):
            return
            
        selected_filter = self.color_filter_combo.currentText()
        
        # D'abord, appliquer le filtre texte pour réinitialiser l'état
        search_text = self.invoice_search_edit.text().lower()
        
        for row in range(self.invoice_table.rowCount()):
            # Vérifier d'abord le filtre texte
            text_match = True
            if search_text:
                text_match = False
                for col in range(self.invoice_table.columnCount()):
                    item = self.invoice_table.item(row, col)
                    if item and search_text in item.text().lower():
                        text_match = True
                        break
            
            # Si le texte ne correspond pas, cacher la ligne
            if not text_match:
                self.invoice_table.setRowHidden(row, True)
                continue
            
            # Vérifier le filtre par couleur
            color_match = True
            if selected_filter != \"Toutes les couleurs\":
                # Obtenir la couleur de la première cellule de la ligne
                item = self.invoice_table.item(row, 0)
                if item:
                    background_color = item.background().color()
                    
                    if selected_filter == \"Validé (Bleu)\":
                        # Couleur bleue (#0078d7)
                        color_match = background_color.name().lower() == \"#0078d7\"
                    elif selected_filter == \"Correspondance exacte (Vert)\":
                        # Couleur verte - utilisons une approche plus simple
                        color_match = (background_color.red() < 150 and 
                                     background_color.green() > 150 and 
                                     background_color.blue() < 150)
                    elif selected_filter == \"Correspondance partielle (Orange)\":
                        # Couleur orange
                        color_match = (background_color.red() > 200 and 
                                     background_color.green() > 100 and 
                                     background_color.green() < 200 and
                                     background_color.blue() < 100)
                    elif selected_filter == \"Pas de correspondance (Sans couleur)\":
                        # Pas de couleur de fond (blanc ou transparent)
                        color_name = background_color.name().lower()
                        color_match = (color_name in [\"#ffffff\", \"#f0f0f0\"] or 
                                     background_color.alpha() == 0 or
                                     (background_color.red() > 240 and 
                                      background_color.green() > 240 and 
                                      background_color.blue() > 240))
                else:
                    color_match = (selected_filter == \"Pas de correspondance (Sans couleur)\")
            
            # Afficher/cacher la ligne selon les deux filtres
            self.invoice_table.setRowHidden(row, not color_match)"""
        
        if old_color_filter in content:
            content = content.replace(old_color_filter, new_color_filter)
            print("[OK] Méthode filter_invoices_by_color corrigée")
        else:
            print("[ATTENTION] Ancienne méthode filter_invoices_by_color non trouvée")
            return False
        
        # Écrire le contenu modifié
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"[OK] Filtre par couleur corrigé avec succès!")
        print(f"[OK] Sauvegarde créée: {backup_file}")
        return True
        
    except Exception as e:
        print(f"Erreur lors de la correction: {str(e)}")
        # Restaurer la sauvegarde en cas d'erreur
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
            print(f"[OK] Fichier restauré depuis la sauvegarde")
        return False

if __name__ == "__main__":
    print("=== Correction du filtre par couleur ===")
    print()
    print("Problèmes à corriger:")
    print("- Logique de filtrage incorrecte")
    print("- Nom de champ de recherche incorrect")
    print("- Détection des couleurs améliorée")
    print()
    
    # Appliquer la correction
    if fix_color_filter():
        print()
        print("=== Résumé ===")
        print("Le filtre par couleur a été corrigé.")
        print("Les factures devraient maintenant s'afficher correctement")
        print("selon le filtre sélectionné.")
        print()
        print("Relancez l'application pour tester le filtre corrigé!")
    else:
        print("Échec de la correction. Vérifiez les messages d'erreur ci-dessus.")
