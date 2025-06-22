"""
FONCTION SIMPLIFIÉE POUR LA SAISIE DES CODES DANS LES FICHIERS EXCEL DE FACTURATION

Cette version simplifiée se base sur le principe qu'il n'y a qu'un numéro de facture par UH,
ce qui rend la correspondance beaucoup plus précise.

Logique de la fonction:
1. Pour chaque ligne validée, aller dans la feuille de calcul qui porte le nom de l'UH
2. Chercher le numéro de facture dans cette feuille spécifique
3. Appliquer les codes client et chorus uniquement pour cette facture
"""

def save_invoice_file(self):
    """Saisie des codes dans le fichier de facturation pour les lignes validées (en bleu)"""
    try:
        # Vérifier si un fichier Excel est chargé
        if not hasattr(self, 'current_excel_file') or not self.current_excel_file:
            # Si aucun fichier n'est chargé, utiliser le dernier fichier ouvert
            if hasattr(self, 'current_invoice_path') and self.current_invoice_path and os.path.exists(self.current_invoice_path):
                self.current_excel_file = self.current_invoice_path
                logger.info(f"Utilisation du dernier fichier ouvert: {self.current_excel_file}")
            else:
                QMessageBox.warning(self, "Attention", "Aucun fichier de facturation n'est chargé.")
                return False
        
        # Collecter les lignes validées (en bleu) du tableau
        validated_rows = []
        for row in range(self.invoice_table.rowCount()):
            # Vérifier si la ligne est validée (en bleu)
            item = self.invoice_table.item(row, 0)  # Vérifier la première cellule
            if item and item.background().color().rgb() == QColor(173, 216, 230).rgb():
                # Récupérer les informations de la ligne
                row_data = {
                    'uh': self.invoice_table.item(row, 0).text() if self.invoice_table.item(row, 0) else "",
                    'facture_num': self.invoice_table.item(row, 1).text() if self.invoice_table.item(row, 1) else "",
                    'nom_facture': self.invoice_table.item(row, 2).text() if self.invoice_table.item(row, 2) else "",
                    'adresse_facture': self.invoice_table.item(row, 3).text() if self.invoice_table.item(row, 3) else "",
                    'nom_bdd': self.invoice_table.item(row, 4).text() if self.invoice_table.item(row, 4) else "",
                    'code_client': self.invoice_table.item(row, 5).text() if self.invoice_table.item(row, 5) else "",
                    'code_chorus': self.invoice_table.item(row, 6).text() if self.invoice_table.item(row, 6) else "",
                    'ligne_bdd': self.invoice_table.item(row, 7).text() if self.invoice_table.item(row, 7) else ""
                }
                validated_rows.append(row_data)
        
        if not validated_rows:
            QMessageBox.warning(self, "Attention", "Aucune ligne validée (en bleu) n'a été trouvée.")
            return False
        
        # Créer un nouveau nom de fichier
        file_name, file_ext = os.path.splitext(self.current_excel_file)
        new_file_path = f"{file_name}_updated{file_ext}"
        
        # Copier le fichier original en utilisant shutil.copy2 qui préserve les métadonnées
        try:
            shutil.copy2(self.current_excel_file, new_file_path)
            logger.info(f"Fichier copié avec succès: {new_file_path}")
        except Exception as e:
            logger.error(f"Erreur lors de la copie du fichier: {str(e)}")
            raise Exception(f"Impossible de copier le fichier Excel: {str(e)}")
        
        # Utiliser win32com pour manipuler Excel directement (plus fiable pour les fichiers complexes)
        try:
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False  # Ne pas afficher Excel
            excel.DisplayAlerts = False  # Désactiver les alertes
            
            # Ouvrir le fichier copié
            workbook = excel.Workbooks.Open(os.path.abspath(new_file_path))
            
            # Compteurs pour le rapport
            factures_traitees = 0
            
            # Pour chaque ligne validée, chercher la feuille correspondant à l'UH
            for row_data in validated_rows:
                uh = row_data['uh']
                facture_num = row_data['facture_num']
                code_client = row_data['code_client']
                code_chorus = row_data['code_chorus']
                
                # Log pour débogage
                logger.info(f"Traitement de la facture {facture_num} (UH: {uh})")
                
                # Facture trouvée et traitée pour cette ligne?
                facture_traitee = False
                
                # Chercher une feuille qui correspond exactement à l'UH
                feuille_uh_trouvee = False
                for i in range(1, workbook.Sheets.Count + 1):
                    sheet = workbook.Sheets(i)
                    sheet_name = sheet.Name
                    
                    # Vérifier si cette feuille correspond à l'UH
                    if uh.lower() in sheet_name.lower():
                        feuille_uh_trouvee = True
                        logger.info(f"Feuille correspondant à l'UH {uh} trouvée: {sheet_name}")
                        
                        # Chercher le numéro de facture et la cellule "Intitulé" dans cette feuille spécifique
                        facture_trouvee = False
                        intitule_cell = None
                        
                        # Rechercher dans toutes les cellules de la feuille
                        for row in range(1, 100):  # Limiter la recherche aux 100 premières lignes
                            for col in range(1, 20):  # Limiter aux 20 premières colonnes
                                try:
                                    cell_value = sheet.Cells(row, col).Value
                                    if cell_value is not None:  # Vérifier que la cellule n'est pas vide
                                        # Convertir en string, quelle que soit la valeur
                                        cell_text = str(cell_value).lower().strip()
                                        
                                        # Chercher le numéro de facture
                                        facture_num_clean = facture_num.strip().lower()
                                        if (facture_num_clean == cell_text or
                                            facture_num_clean in cell_text or
                                            f"facture {facture_num_clean}" in cell_text or
                                            f"facture n°{facture_num_clean}" in cell_text or
                                            f"facture n° {facture_num_clean}" in cell_text):
                                            facture_trouvee = True
                                            logger.info(f"Numéro de facture {facture_num} trouvé dans la cellule ({row}, {col}): '{cell_text}'")
                                        
                                        # Chercher la cellule "Intitulé"
                                        if "intitulé" in cell_text or "intitule" in cell_text:
                                            intitule_cell = (row, col)
                                            logger.info(f"Cellule 'Intitulé' trouvée en ({row}, {col})")
                                except Exception as e:
                                    logger.debug(f"Erreur lors de la lecture de la cellule ({row}, {col}): {str(e)}")
                                    continue
                        
                        # Si on a trouvé la cellule "Intitulé", on peut insérer les codes
                        # Même si on n'a pas trouvé le numéro de facture exact, on sait qu'il n'y a qu'une facture par UH
                        if intitule_cell:
                            intitule_row, intitule_col = intitule_cell
                            
                            # Insérer le code client: 2 cellules plus bas et une cellule à droite de "Intitulé"
                            if code_client:
                                client_cell_row = intitule_row + 2
                                client_cell_col = intitule_col + 1
                                sheet.Cells(client_cell_row, client_cell_col).Value = code_client
                                logger.info(f"Code client {code_client} inséré dans la cellule ({client_cell_row}, {client_cell_col})")
                            
                            # Insérer le code chorus: une cellule en bas et une cellule à droite de "Intitulé"
                            if code_chorus:
                                chorus_cell_row = intitule_row + 1
                                chorus_cell_col = intitule_col + 1
                                sheet.Cells(chorus_cell_row, chorus_cell_col).Value = code_chorus
                                logger.info(f"Code chorus {code_chorus} inséré dans la cellule ({chorus_cell_row}, {chorus_cell_col})")
                            
                            factures_traitees += 1
                            facture_traitee = True
                            
                            if facture_trouvee:
                                logger.info(f"Codes insérés pour la facture {facture_num} dans la feuille {sheet_name}")
                            else:
                                logger.info(f"Codes insérés dans la feuille {sheet_name} correspondant à l'UH {uh} (numéro de facture non trouvé explicitement)")
                            
                            break  # Sortir de la boucle des feuilles pour cette ligne
                
                if not feuille_uh_trouvee:
                    logger.warning(f"Aucune feuille correspondant à l'UH {uh} n'a été trouvée")
            
            # Sauvegarder et fermer le fichier
            workbook.Save()
            workbook.Close(True)
            excel.Quit()
            
            # Libérer les ressources COM
            del workbook
            del excel
            
        except ImportError:
            # Si win32com n'est pas disponible, utiliser openpyxl comme fallback
            logger.warning("win32com n'est pas disponible, utilisation d'openpyxl comme alternative")
            
            # Ouvrir le fichier avec openpyxl
            workbook = load_workbook(new_file_path)
            
            # Compteurs pour le rapport
            factures_traitees = 0
            
            # Pour chaque ligne validée, chercher la feuille correspondant à l'UH
            for row_data in validated_rows:
                uh = row_data['uh']
                facture_num = row_data['facture_num']
                code_client = row_data['code_client']
                code_chorus = row_data['code_chorus']
                
                # Log pour débogage
                logger.info(f"Traitement de la facture {facture_num} (UH: {uh})")
                
                # Facture trouvée et traitée pour cette ligne?
                facture_traitee = False
                
                # Chercher une feuille qui correspond exactement à l'UH
                feuille_uh_trouvee = False
                for sheet_name in workbook.sheetnames:
                    # Vérifier si cette feuille correspond à l'UH
                    if uh.lower() in sheet_name.lower():
                        feuille_uh_trouvee = True
                        sheet = workbook[sheet_name]
                        logger.info(f"Feuille correspondant à l'UH {uh} trouvée: {sheet_name}")
                        
                        # Chercher le numéro de facture et la cellule "Intitulé" dans cette feuille spécifique
                        facture_trouvee = False
                        intitule_cell = None
                        
                        # Rechercher dans toutes les cellules de la feuille
                        for row in range(1, min(100, sheet.max_row + 1)):  # Limiter la recherche aux 100 premières lignes
                            for col in range(1, min(20, sheet.max_column + 1)):  # Limiter aux 20 premières colonnes
                                try:
                                    cell_value = sheet.cell(row=row, column=col).value
                                    if cell_value is not None:  # Vérifier que la cellule n'est pas vide
                                        # Convertir en string, quelle que soit la valeur
                                        cell_text = str(cell_value).lower().strip()
                                        
                                        # Chercher le numéro de facture
                                        facture_num_clean = facture_num.strip().lower()
                                        if (facture_num_clean == cell_text or
                                            facture_num_clean in cell_text or
                                            f"facture {facture_num_clean}" in cell_text or
                                            f"facture n°{facture_num_clean}" in cell_text or
                                            f"facture n° {facture_num_clean}" in cell_text):
                                            facture_trouvee = True
                                            logger.info(f"Numéro de facture {facture_num} trouvé dans la cellule ({row}, {col}): '{cell_text}'")
                                        
                                        # Chercher la cellule "Intitulé"
                                        if "intitulé" in cell_text or "intitule" in cell_text:
                                            intitule_cell = (row, col)
                                            logger.info(f"Cellule 'Intitulé' trouvée en ({row}, {col})")
                                except Exception as e:
                                    logger.debug(f"Erreur lors de la lecture de la cellule ({row}, {col}): {str(e)}")
                                    continue
                        
                        # Si on a trouvé la cellule "Intitulé", on peut insérer les codes
                        # Même si on n'a pas trouvé le numéro de facture exact, on sait qu'il n'y a qu'une facture par UH
                        if intitule_cell:
                            intitule_row, intitule_col = intitule_cell
                            
                            # Insérer le code client: 2 cellules plus bas et une cellule à droite de "Intitulé"
                            if code_client:
                                client_cell_row = intitule_row + 2
                                client_cell_col = intitule_col + 1
                                sheet.cell(row=client_cell_row, column=client_cell_col).value = code_client
                                logger.info(f"Code client {code_client} inséré dans la cellule ({client_cell_row}, {client_cell_col})")
                            
                            # Insérer le code chorus: une cellule en bas et une cellule à droite de "Intitulé"
                            if code_chorus:
                                chorus_cell_row = intitule_row + 1
                                chorus_cell_col = intitule_col + 1
                                sheet.cell(row=chorus_cell_row, column=chorus_cell_col).value = code_chorus
                                logger.info(f"Code chorus {code_chorus} inséré dans la cellule ({chorus_cell_row}, {chorus_cell_col})")
                            
                            factures_traitees += 1
                            facture_traitee = True
                            
                            if facture_trouvee:
                                logger.info(f"Codes insérés pour la facture {facture_num} dans la feuille {sheet_name}")
                            else:
                                logger.info(f"Codes insérés dans la feuille {sheet_name} correspondant à l'UH {uh} (numéro de facture non trouvé explicitement)")
                            
                            break  # Sortir de la boucle des feuilles pour cette ligne
                
                if not feuille_uh_trouvee:
                    logger.warning(f"Aucune feuille correspondant à l'UH {uh} n'a été trouvée")
            
            # Sauvegarder et fermer le fichier
            workbook.save(new_file_path)
            workbook.close()
        
        except Exception as e:
            logger.error(f"Erreur lors de la manipulation du fichier Excel: {str(e)}")
            raise Exception(f"Erreur lors de la manipulation du fichier Excel: {str(e)}")
        
        # Vérifier que le fichier a été correctement créé
        if not os.path.exists(new_file_path) or os.path.getsize(new_file_path) == 0:
            raise Exception("Erreur lors de la création du fichier Excel. Le fichier est vide ou n'a pas été créé.")
        
        if factures_traitees > 0:
            # Informer l'utilisateur du succès
            QMessageBox.information(
                self,
                "Succès",
                f"Les codes ont été insérés pour {factures_traitees} facture(s) dans le fichier:\n{new_file_path}"
            )
            
            # Proposer d'ouvrir le fichier Excel
            reply = QMessageBox.question(
                self,
                "Ouvrir le fichier",
                "Voulez-vous ouvrir le fichier Excel généré ?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            
            if reply == QMessageBox.Yes:
                try:
                    # Ouvrir le fichier avec l'application par défaut
                    os.startfile(os.path.abspath(new_file_path))
                    logger.info(f"Fichier ouvert: {new_file_path}")
                except Exception as e:
                    logger.error(f"Erreur lors de l'ouverture du fichier: {str(e)}")
                    QMessageBox.warning(self, "Attention", f"Impossible d'ouvrir le fichier: {str(e)}")
        else:
            QMessageBox.warning(
                self,
                "Attention",
                f"Aucune facture n'a été trouvée dans le fichier Excel. Vérifiez que les numéros de facture correspondent."
            )
        
        return True
        
    except Exception as e:
        logger.error(f"Erreur lors de la saisie des codes: {str(e)}", exc_info=True)
        QMessageBox.critical(
            self,
            "Erreur",
            f"Une erreur est survenue lors de la saisie des codes:\n{str(e)}"
        )
        return False
