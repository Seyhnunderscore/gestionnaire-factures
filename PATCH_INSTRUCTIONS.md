# INSTRUCTIONS POUR CORRIGER LE PROBLÈME DE MAPPING DES BOUTONS VALIDÉR

## Problème identifié :
- Les boutons "Valider" restent connectés aux anciens index après tri du tableau
- Le tableau a 8 colonnes mais le bouton est placé en colonne 8 (qui n'existe pas)
- Cela cause le dysfonctionnement décrit par l'utilisateur

## Corrections à apporter dans main.py :

### 1. Ligne 2258 - Corriger le nombre de colonnes :
```python
# REMPLACER :
self.invoice_table.setColumnCount(8)  # 8 colonnes pour les données de facture

# PAR :
self.invoice_table.setColumnCount(9)  # 9 colonnes pour inclure le bouton Statut
```

### 2. Ligne 2259 - Corriger les en-têtes de colonnes :
```python
# REMPLACER :
self.invoice_table.setHorizontalHeaderLabels(["UH", "N° Facture", "Nom facture", "Adresse facture", "Nom BDD", "Date", "Montant", "Statut"])

# PAR :
self.invoice_table.setHorizontalHeaderLabels(["UH", "N° Facture", "Nom facture", "Adresse facture", "Nom BDD", "Date", "Montant", "Ligne BDD", "Statut"])
```

### 3. Ligne 2315 - Corriger la connexion du bouton :
```python
# REMPLACER :
validate_btn.clicked.connect(lambda checked, row=i: self.validate_invoice_row(row))

# PAR :
validate_btn.clicked.connect(self.on_validate_button_clicked)
```

### 4. Ajouter la nouvelle méthode on_validate_button_clicked APRÈS la méthode update_invoice_table :

```python
def on_validate_button_clicked(self):
    """Méthode appelée quand un bouton Valider est cliqué - retrouve dynamiquement la ligne"""
    try:
        # Récupérer le bouton qui a été cliqué
        button = self.sender()
        if not button:
            return
        
        # Parcourir toutes les cellules du tableau pour trouver celle qui contient ce bouton
        current_row = -1
        for row in range(self.invoice_table.rowCount()):
            widget = self.invoice_table.cellWidget(row, 8)  # Colonne Statut
            if widget == button:
                current_row = row
                break
        
        # Si on a trouvé la ligne, valider la facture
        if current_row >= 0:
            self.validate_invoice_row_safe(current_row)
        else:
            logger.warning("Impossible de trouver la ligne correspondant au bouton cliqué")
            
    except Exception as e:
        logger.error(f"Erreur lors du clic sur le bouton Valider: {str(e)}", exc_info=True)

def validate_invoice_row_safe(self, visual_row):
    """Valide une ligne de facture en utilisant la position visuelle actuelle"""
    try:
        # Vérifier si la ligne existe
        if visual_row < 0 or visual_row >= self.invoice_table.rowCount():
            logger.warning(f"Tentative de validation d'une ligne invalide: {visual_row}")
            return
        
        # Récupérer les données de la ligne visuelle actuelle
        uh_item = self.invoice_table.item(visual_row, 0)
        numero_item = self.invoice_table.item(visual_row, 1)
        client_item = self.invoice_table.item(visual_row, 2)
        
        if not uh_item or not numero_item:
            logger.warning(f"Données manquantes pour la ligne {visual_row}")
            return
        
        uh = uh_item.text()
        numero = numero_item.text()
        client = client_item.text() if client_item else ""
        
        logger.info(f"Validation de la facture ligne {visual_row}: UH={uh}, N°={numero}")
        
        # Changer la couleur de la ligne pour indiquer qu'elle est validée
        color = QColor("#0078d7")  # Bleu pour "Validé"
        
        # Appliquer la couleur à toutes les cellules de la ligne
        for col in range(self.invoice_table.columnCount()):
            item = self.invoice_table.item(visual_row, col)
            if item:
                item.setBackground(color)
        
        # Mettre à jour le bouton
        button = self.invoice_table.cellWidget(visual_row, 8)
        if isinstance(button, QPushButton):
            button.setStyleSheet("background-color: #87CEFA; color: white; font-weight: bold;")
            button.setText("Validé")
        
        # Mettre à jour les statistiques
        if hasattr(self, 'update_statistics'):
            self.update_statistics()
        
        logger.info(f"Facture ligne {visual_row} validée avec succès")
        
    except Exception as e:
        logger.error(f"Erreur lors de la validation de la facture ligne {visual_row}: {str(e)}", exc_info=True)
        QMessageBox.critical(
            self,
            "Erreur",
            f"Une erreur est survenue lors de la validation de la facture:\n{str(e)}"
        )
```

### 5. Corrections dans la boucle de création des lignes (vers ligne 2304-2310) :

```python
# Corriger l'assignation de la colonne "Ligne BDD" :
# REMPLACER ligne 2310 :
self.invoice_table.setItem(i, 7, ligne_bdd_item)  # Ligne BDD (éditable)

# ET REMPLACER ligne 2316 :
self.invoice_table.setCellWidget(i, 8, validate_btn)  # Statut (bouton Valider)
```

## Résumé des corrections :
1. **Nombre de colonnes** : 8 → 9
2. **En-têtes** : Ajout de "Ligne BDD" comme colonne séparée
3. **Connexion boutons** : Utilisation d'une méthode dynamique au lieu d'un index fixe
4. **Nouvelles méthodes** : `on_validate_button_clicked()` et `validate_invoice_row_safe()`

Ces corrections résoudront le problème de désynchronisation entre les boutons et les lignes après tri du tableau.
