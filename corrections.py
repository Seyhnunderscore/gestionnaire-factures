"""
Ce fichier contient les corrections à apporter au fichier main.py
"""

# CORRECTION 1: Méthode load_state_async modifiée pour éviter l'ouverture automatique de l'explorateur Windows
def load_state_async(self):
    """Charge l'état de l'application de manière asynchrone"""
    try:
        # Cette méthode est appelée de manière asynchrone après le chargement de l'interface
        # pour éviter de bloquer l'interface utilisateur
        
        # Vérifier si un fichier de facture récent doit être chargé
        # Commenté pour éviter l'ouverture automatique de l'explorateur Windows au démarrage
        # if hasattr(self, 'current_invoice_path') and self.current_invoice_path:
        #     if os.path.exists(self.current_invoice_path):
        #         # Charger le fichier de facture de manière asynchrone
        #         QTimer.singleShot(100, lambda: self.load_invoice_file())
        
        # Autres opérations de chargement asynchrone peuvent être ajoutées ici
        
        logger.info("Chargement asynchrone de l'état terminé")
        
    except Exception as e:
        logger.error(f"Erreur lors du chargement asynchrone de l'état: {str(e)}")

# CORRECTION 2: Méthode process_invoice_file modifiée pour mieux extraire les factures
def process_invoice_file(self, file_path):
    """Traite un fichier Excel de facturation pour en extraire les informations
    
    Args:
        file_path (str): Chemin vers le fichier Excel à traiter
    """
    try:
        # Afficher un message de progression
        progress_dialog = QProgressDialog("Traitement du fichier de facturation...", "Annuler", 0, 100, self)
        progress_dialog.setWindowTitle("Importation des factures")
        progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        progress_dialog.setMinimumDuration(0)
        progress_dialog.show()
        QApplication.processEvents()
        
        # Charger le fichier Excel avec openpyxl
        workbook = load_workbook(file_path, data_only=True)
        
        # Initialiser la liste des factures
        self.invoices = []
        
        # Nombre total de feuilles pour le calcul de progression
        total_sheets = len(workbook.sheetnames)
        processed_sheets = 0
        
        # Parcourir chaque feuille (UH)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            
            # Mettre à jour la progression
            processed_sheets += 1
            progress_percent = int((processed_sheets / total_sheets) * 100)
            progress_dialog.setValue(progress_percent)
            progress_dialog.setLabelText(f"Traitement de l'UH: {sheet_name}")
            QApplication.processEvents()
            
            if progress_dialog.wasCanceled():
                break
            
            # Rechercher les cellules contenant "Intitulé"
            for row_idx in range(1, sheet.max_row + 1):
                for col_idx in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row=row_idx, column=col_idx)
                    
                    # Si la cellule contient "Intitulé", c'est le début d'une facture
                    if cell.value and "Intitulé" in str(cell.value):
                        try:
                            # Extraire le nom de la facture (6 cellules fusionnées à droite)
                            invoice_name_cell = sheet.cell(row=row_idx, column=col_idx + 1)
                            invoice_name = invoice_name_cell.value if invoice_name_cell.value else ""
                            
                            # Extraire le numéro de facture (7 cellules à droite)
                            invoice_number_cell = sheet.cell(row=row_idx, column=col_idx + 7)
                            invoice_number = invoice_number_cell.value if invoice_number_cell.value else ""
                            
                            # Extraire l'adresse (2 cellules en bas et 5 cellules à droite)
                            address_cell = sheet.cell(row=row_idx + 2, column=col_idx + 5)
                            address = address_cell.value if address_cell.value else ""
                            
                            # Vérifier si les cellules sont fusionnées et extraire toutes les valeurs
                            # Pour le nom de la facture (qui peut être sur plusieurs cellules)
                            for merged_cell in sheet.merged_cells.ranges:
                                if invoice_name_cell.coordinate in merged_cell:
                                    # Extraire toutes les valeurs des cellules fusionnées
                                    for row in range(merged_cell.min_row, merged_cell.max_row + 1):
                                        for col in range(merged_cell.min_col, merged_cell.max_col + 1):
                                            cell_value = sheet.cell(row=row, column=col).value
                                            if cell_value and not invoice_name:
                                                invoice_name = cell_value
                            
                            # Pour l'adresse (qui peut être sur plusieurs cellules)
                            for merged_cell in sheet.merged_cells.ranges:
                                if address_cell.coordinate in merged_cell:
                                    # Extraire toutes les valeurs des cellules fusionnées
                                    for row in range(merged_cell.min_row, merged_cell.max_row + 1):
                                        for col in range(merged_cell.min_col, merged_cell.max_col + 1):
                                            cell_value = sheet.cell(row=row, column=col).value
                                            if cell_value and not address:
                                                address = cell_value
                            
                            # Créer un dictionnaire pour la facture
                            invoice = {
                                "uh": sheet_name,
                                "client": invoice_name,
                                "numero": invoice_number,
                                "adresse": address,
                                "date": datetime.now().strftime("%Y-%m-%d"),  # Date par défaut
                                "montant": 0.0,  # Montant par défaut
                                "statut": "Importée"  # Statut par défaut
                            }
                            
                            # Ajouter la facture à la liste
                            self.invoices.append(invoice)
                            
                            # Log pour débogage
                            logger.debug(f"Facture trouvée: {invoice}")
                        except Exception as e:
                            logger.error(f"Erreur lors de l'extraction d'une facture: {str(e)}")
        
        # Fermer la boîte de dialogue de progression
        progress_dialog.close()
        
        # Mettre à jour le tableau des factures
        self.update_invoice_table()
        
        # Afficher un message de succès
        QMessageBox.information(
            self,
            "Importation terminée",
            f"{len(self.invoices)} factures ont été importées avec succès."
        )
        
    except Exception as e:
        logger.error(f"Erreur lors du traitement du fichier de facturation: {str(e)}", exc_info=True)
        QMessageBox.critical(
            self,
            "Erreur de traitement",
            f"Une erreur est survenue lors du traitement du fichier: {str(e)}"
        )

# CORRECTION 3: Méthode load_invoice_file modifiée pour s'assurer qu'elle fonctionne correctement
def load_invoice_file(self):
    """Charge un fichier de facturation"""
    try:
        # Afficher la boîte de dialogue pour sélectionner le fichier
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Sélectionner un fichier Excel", 
            "", 
            "Fichiers Excel (*.xlsx *.xls *.xlsm);;Tous les fichiers (*)"
        )
        
        if not file_path:
            return
            
        # Afficher un indicateur de chargement
        self.statusBar().showMessage("Chargement du fichier en cours...")
        QApplication.processEvents()  # Mettre à jour l'interface
        
        # Vérifier que le fichier existe et est accessible
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"Le fichier spécifié n'existe pas: {file_path}")
            
        if not os.access(file_path, os.R_OK):
            raise PermissionError(f"Impossible de lire le fichier. Vérifiez les permissions: {file_path}")
        
        # Sauvegarder le chemin du fichier
        self.current_excel_file = file_path
        
        # Traiter le fichier
        self.process_invoice_file(file_path)
        
        # Activer le bouton d'ouverture avec Excel si nécessaire
        if hasattr(self, 'open_in_excel_btn'):
            self.open_in_excel_btn.setEnabled(True)
        
        # Mettre à jour le titre de la fenêtre avec le nom du fichier
        self.setWindowTitle(f"Gestionnaire de Factures - {os.path.basename(file_path)}")
        self.statusBar().showMessage(f"Fichier chargé avec succès: {os.path.basename(file_path)}", 5000)
        
    except FileNotFoundError as e:
        error_msg = f"Fichier introuvable : {str(e)}"
        logger.error(error_msg)
        QMessageBox.critical(self, "Erreur de fichier", error_msg)
        
    except PermissionError as e:
        error_msg = f"Erreur de permission : {str(e)}"
        logger.error(error_msg)
        QMessageBox.critical(self, "Erreur d'accès", error_msg)
        
    except Exception as e:
        error_msg = f"Erreur lors du chargement du fichier : {str(e)}"
        logger.error(error_msg, exc_info=True)
        QMessageBox.critical(self, "Erreur", error_msg)
        
    finally:
        # S'assurer que la barre d'état est réinitialisée
        self.statusBar().clearMessage()
