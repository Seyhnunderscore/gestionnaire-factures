#!/usr/bin/env python3
"""
Script pour ajouter un filtre par couleur dans la partie Factures
"""

import os
import shutil
from datetime import datetime

def add_color_filter():
    """Ajoute un filtre par couleur dans l'interface des factures"""
    
    # Chemin vers le fichier main.py
    main_file = "main.py"
    backup_file = f"main.py.backup_color_filter_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
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
        
        # 1. Ajouter le filtre par couleur après la barre de recherche
        search_section = """        self.invoice_layout.addLayout(search_layout)"""
        
        color_filter_code = """        self.invoice_layout.addLayout(search_layout)
        
        # Filtre par couleur
        color_filter_layout = QHBoxLayout()
        color_filter_layout.addWidget(QLabel("Filtrer par couleur:"))
        
        self.color_filter_combo = QComboBox()
        self.color_filter_combo.addItem("Toutes les couleurs")
        self.color_filter_combo.addItem("Validé (Bleu)")
        self.color_filter_combo.addItem("Correspondance exacte (Vert)")
        self.color_filter_combo.addItem("Correspondance partielle (Orange)")
        self.color_filter_combo.addItem("Pas de correspondance (Sans couleur)")
        self.color_filter_combo.currentTextChanged.connect(self.filter_invoices_by_color)
        color_filter_layout.addWidget(self.color_filter_combo)
        color_filter_layout.addStretch()
        
        self.invoice_layout.addLayout(color_filter_layout)"""
        
        if search_section in content:
            content = content.replace(search_section, color_filter_code)
            print("[OK] Filtre par couleur ajouté dans l'interface")
        else:
            print("[ATTENTION] Section de recherche non trouvée")
            return False
        
        # 2. Ajouter la méthode filter_invoices_by_color
        # Trouver la fin de la classe MainWindow pour ajouter la nouvelle méthode
        filter_invoices_method = """    def filter_invoices(self):
        \"\"\"Filtre les factures en fonction de la recherche\"\"\"
        search_text = self.invoice_search_input.text().lower()
        
        for row in range(self.invoice_table.rowCount()):
            hide_row = True
            for col in range(self.invoice_table.columnCount()):
                item = self.invoice_table.item(row, col)
                if item and search_text in item.text().lower():
                    hide_row = False
                    break
            
            self.invoice_table.setRowHidden(row, hide_row)"""
        
        new_filter_method = """    def filter_invoices(self):
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
        self.filter_invoices_by_color()
    
    def filter_invoices_by_color(self):
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
            
            if selected_filter == "Toutes les couleurs":
                hide_row = False
            elif selected_filter == "Validé (Bleu)":
                # Couleur bleue (#0078d7)
                hide_row = background_color.name().lower() != "#0078d7"
            elif selected_filter == "Correspondance exacte (Vert)":
                # Couleur verte (success color)
                hide_row = not (background_color.red() < 100 and background_color.green() > 150 and background_color.blue() < 100)
            elif selected_filter == "Correspondance partielle (Orange)":
                # Couleur orange (warning color)
                hide_row = not (background_color.red() > 200 and background_color.green() > 100 and background_color.blue() < 100)
            elif selected_filter == "Pas de correspondance (Sans couleur)":
                # Pas de couleur de fond ou couleur par défaut
                hide_row = background_color.name().lower() not in ["#ffffff", "#f0f0f0", "#transparent"]
            
            self.invoice_table.setRowHidden(row, hide_row)"""
        
        if filter_invoices_method in content:
            content = content.replace(filter_invoices_method, new_filter_method)
            print("[OK] Méthode filter_invoices_by_color ajoutée")
        else:
            print("[ATTENTION] Méthode filter_invoices non trouvée")
            return False
        
        # Écrire le contenu modifié
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print(f"[OK] Filtre par couleur ajouté avec succès!")
        print(f"[OK] Sauvegarde créée: {backup_file}")
        return True
        
    except Exception as e:
        print(f"Erreur lors de l'ajout du filtre: {str(e)}")
        # Restaurer la sauvegarde en cas d'erreur
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
            print(f"[OK] Fichier restauré depuis la sauvegarde")
        return False

def verify_color_filter():
    """Vérifie que le filtre par couleur a été ajouté"""
    main_file = "main.py"
    
    if not os.path.exists(main_file):
        print("Erreur: Impossible de vérifier, le fichier main.py n'existe pas")
        return False
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Vérifier la présence du filtre
        if "color_filter_combo" in content and "filter_invoices_by_color" in content:
            print("[OK] Vérification réussie: le filtre par couleur est présent")
            return True
        else:
            print("[ATTENTION] Le filtre par couleur n'a pas été ajouté correctement")
            return False
            
    except Exception as e:
        print(f"Erreur lors de la vérification: {str(e)}")
        return False

if __name__ == "__main__":
    print("=== Ajout du filtre par couleur dans les Factures ===")
    print()
    print("Fonctionnalité à ajouter:")
    print("- Menu déroulant pour filtrer par couleur")
    print("- Options: Toutes, Validé (Bleu), Exacte (Vert), Partielle (Orange), Sans correspondance")
    print("- Intégration avec le filtre de recherche existant")
    print()
    
    # Appliquer l'ajout
    if add_color_filter():
        print()
        print("=== Vérification de l'ajout ===")
        verify_color_filter()
        print()
        print("=== Résumé ===")
        print("Le filtre par couleur a été ajouté à l'interface des factures.")
        print("Vous pouvez maintenant filtrer les factures par:")
        print("- Toutes les couleurs")
        print("- Validé (Bleu)")
        print("- Correspondance exacte (Vert)")
        print("- Correspondance partielle (Orange)")
        print("- Pas de correspondance (Sans couleur)")
        print()
        print("Relancez l'application pour voir le nouveau filtre!")
    else:
        print("Échec de l'ajout. Vérifiez les messages d'erreur ci-dessus.")
