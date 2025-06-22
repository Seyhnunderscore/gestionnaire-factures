# Solution pour le problème de mapping des boutons après tri du tableau

def update_invoice_table_fixed(self):
    """Met à jour le tableau des factures avec les données importées - VERSION CORRIGÉE"""
    try:
        # Vérifier si le tableau des factures existe
        if not hasattr(self, 'invoice_table'):
            logger.warning("Le tableau des factures n'existe pas encore, création...")
            # Créer le tableau des factures s'il n'existe pas
            self.invoice_table = QTableWidget()
            # CORRECTION 1: Passer à 9 colonnes pour inclure le bouton
            self.invoice_table.setColumnCount(9)  # 9 colonnes pour inclure le bouton Statut
            self.invoice_table.setHorizontalHeaderLabels([
                "UH", "N° Facture", "Nom facture", "Adresse facture", 
                "Nom BDD", "Date", "Montant", "Ligne BDD", "Statut"
            ])
            self.invoice_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            
            # Ajouter le tableau à l'interface (code existant conservé)
            if hasattr(self, 'invoice_tab'):
                layout = self.invoice_tab.layout()
                if layout:
                    layout.addWidget(self.invoice_table)
                else:
                    layout = QVBoxLayout(self.invoice_tab)
                    layout.addWidget(self.invoice_table)
        
        # Effacer le tableau
        self.invoice_table.setRowCount(0)
        
        # Vérifier si la liste des factures existe
        if not hasattr(self, 'invoices') or not self.invoices:
            logger.warning("Aucune facture à afficher")
            return
        
        # Ajouter les factures au tableau
        for i, invoice in enumerate(self.invoices):
            self.invoice_table.insertRow(i)
            
            # Ajouter les données de la facture au tableau
            self.invoice_table.setItem(i, 0, QTableWidgetItem(invoice.get("uh", "")))
            self.invoice_table.setItem(i, 1, QTableWidgetItem(str(invoice.get("numero", ""))))
            self.invoice_table.setItem(i, 2, QTableWidgetItem(invoice.get("client", "")))
            self.invoice_table.setItem(i, 3, QTableWidgetItem(invoice.get("adresse", "")))
            self.invoice_table.setItem(i, 4, QTableWidgetItem(""))  # Nom BDD (vide)
            self.invoice_table.setItem(i, 5, QTableWidgetItem(""))  # Date (vide)
            self.invoice_table.setItem(i, 6, QTableWidgetItem(""))  # Montant (vide)
            
            # Créer un item pour la colonne "Ligne BDD" qui sera éditable
            ligne_bdd_item = QTableWidgetItem("")
            ligne_bdd_item.setFlags(ligne_bdd_item.flags() | Qt.ItemFlag.ItemIsEditable)
            self.invoice_table.setItem(i, 7, ligne_bdd_item)  # CORRECTION 2: Ligne BDD en colonne 7
            
            # CORRECTION 3: Créer un bouton "Valider" avec une connexion dynamique
            validate_btn = QPushButton("Valider")
            validate_btn.setStyleSheet("background-color: #4CAF50; color: white;")
            # Utiliser une méthode qui retrouve dynamiquement la ligne du bouton
            validate_btn.clicked.connect(self.on_validate_button_clicked)
            self.invoice_table.setCellWidget(i, 8, validate_btn)  # CORRECTION 4: Statut en colonne 8
        
        # Ajuster les colonnes pour qu'elles s'adaptent au contenu
        self.invoice_table.resizeColumnsToContents()
        
        logger.info(f"{len(self.invoices)} factures affichées dans le tableau")
        
        # Mettre à jour les statistiques
        self.update_statistics()
        
    except Exception as e:
        logger.error(f"Erreur lors de la mise à jour du tableau des factures: {str(e)}", exc_info=True)


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
        self.update_statistics()
        
        logger.info(f"Facture ligne {visual_row} validée avec succès")
        
    except Exception as e:
        logger.error(f"Erreur lors de la validation de la facture ligne {visual_row}: {str(e)}", exc_info=True)
        QMessageBox.critical(
            self,
            "Erreur",
            f"Une erreur est survenue lors de la validation de la facture:\n{str(e)}"
        )
