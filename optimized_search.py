# Fonction optimisée pour la recherche des factures
def optimized_win32com_search(self, workbook, validated_rows):
    """Version optimisée de la recherche de factures avec win32com"""
    # Compteurs pour le rapport
    factures_traitees = 0
    
    # Pour chaque ligne validée, chercher la feuille correspondant à l'UH
    for row_data in validated_rows:
        uh = row_data['uh']
        facture_num = row_data['facture_num'].strip()
        code_client = row_data['code_client']
        code_chorus = row_data['code_chorus']
        
        # Extraire le numéro pur de la facture (sans préfixe)
        facture_num_pur = None
        if "facture n°" in facture_num.lower():
            facture_num_pur = facture_num.lower().replace("facture n°", "").strip()
        elif "facture n" in facture_num.lower():
            facture_num_pur = facture_num.lower().replace("facture n", "").strip()
        elif "facture" in facture_num.lower():
            facture_num_pur = facture_num.lower().replace("facture", "").strip()
        else:
            facture_num_pur = facture_num.lower().strip()
        
        # Log pour débogage
        logger.info(f"Traitement de la facture {facture_num} (UH: {uh}, numéro pur: {facture_num_pur})")
        
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
                
                # Chercher le numéro de facture exact dans cette feuille spécifique
                facture_trouvee = False
                facture_row = None
                facture_col = None
                
                # Rechercher uniquement dans les colonnes 8 (H) et 16 (P) où se trouvent les numéros de facture
                colonnes_facture = [8, 16]  # Colonnes H et P
                lignes_vides_consecutives = 0
                max_lignes_vides = 20  # Arrêter après 20 lignes vides consécutives
                
                for row in range(1, 2000):  # Limite maximale de 2000 lignes
                    ligne_contient_donnees = False
                    
                    for col in colonnes_facture:  # Uniquement les colonnes où se trouvent les factures
                        try:
                            cell_value = sheet.Cells(row, col).Value
                            if cell_value is not None:  # Vérifier que la cellule n'est pas vide
                                ligne_contient_donnees = True
                                # Convertir en string, quelle que soit la valeur
                                cell_text = str(cell_value).lower().strip()
                                
                                # Extraire le numéro pur de la cellule et le normaliser
                                cell_num_pur = None
                                if "facture n°" in cell_text:
                                    cell_num_pur = cell_text.replace("facture n°", "").strip()
                                elif "facture n" in cell_text:
                                    cell_num_pur = cell_text.replace("facture n", "").strip()
                                elif "facture" in cell_text:
                                    cell_num_pur = cell_text.replace("facture", "").strip()
                                else:
                                    cell_num_pur = cell_text
                                
                                # Normaliser les numéros en supprimant les espaces et caractères non numériques
                                normalized_cell_num = ''.join(c for c in cell_num_pur if c.isdigit())
                                normalized_facture_num = ''.join(c for c in facture_num_pur if c.isdigit())
                                
                                logger.debug(f"Comparaison: '{normalized_cell_num}' vs '{normalized_facture_num}'" +
                                            f" (original: '{cell_text}' vs '{facture_num}')")
                                
                                # Vérification avec plusieurs méthodes
                                if (cell_num_pur == facture_num_pur or  # Méthode 1: Comparaison exacte
                                    normalized_cell_num == normalized_facture_num):  # Méthode 2: Comparaison des chiffres uniquement
                                    facture_trouvee = True
                                    facture_row = row
                                    facture_col = col
                                    logger.info(f"Numéro de facture {facture_num} trouvé dans la cellule ({row}, {col}): '{cell_text}'")
                                    break
                        except Exception as e:
                            logger.debug(f"Erreur lors de la lecture de la cellule ({row}, {col}): {str(e)}")
                            continue
                    
                    # Si on a trouvé la facture, on peut sortir de la boucle des lignes
                    if facture_trouvee:
                        break
                    
                    # Vérifier si la ligne contient des données pour gérer le compteur de lignes vides
                    if not ligne_contient_donnees:
                        lignes_vides_consecutives += 1
                    else:
                        lignes_vides_consecutives = 0  # Réinitialiser le compteur si on trouve des données
                    
                    # Arrêter la recherche après 20 lignes vides consécutives
                    if lignes_vides_consecutives >= max_lignes_vides:
                        logger.info(f"Arrêt de la recherche après {max_lignes_vides} lignes vides consécutives à la ligne {row}")
                        break
                
                # Si on a trouvé le numéro de facture exact, insérer les codes
                if facture_trouvee:
                    # Insérer le code client: 6 cellules à gauche et 2 cellules en dessous du numéro de facture
                    if code_client:
                        client_cell_row = facture_row + 2
                        client_cell_col = facture_col - 6
                        if client_cell_col < 1:  # Vérifier que la colonne est valide
                            client_cell_col = 1
                        
                        # Log avant insertion pour vérification
                        logger.info(f"Position du numéro de facture: ({facture_row}, {facture_col})")
                        logger.info(f"Position calculée pour le code client: ({client_cell_row}, {client_cell_col})")
                        
                        sheet.Cells(client_cell_row, client_cell_col).Value = code_client
                        logger.info(f"Code client {code_client} inséré dans la cellule ({client_cell_row}, {client_cell_col})")
                    
                    # Insérer le code chorus: 6 cellules à gauche et 1 cellule en dessous du numéro de facture
                    if code_chorus:
                        chorus_cell_row = facture_row + 1
                        chorus_cell_col = facture_col - 6
                        if chorus_cell_col < 1:  # Vérifier que la colonne est valide
                            chorus_cell_col = 1
                        
                        # Log avant insertion pour vérification
                        logger.info(f"Position calculée pour le code chorus: ({chorus_cell_row}, {chorus_cell_col})")
                        
                        sheet.Cells(chorus_cell_row, chorus_cell_col).Value = code_chorus
                        logger.info(f"Code chorus {code_chorus} inséré dans la cellule ({chorus_cell_row}, {chorus_cell_col})")
                    
                    facture_traitee = True
                    factures_traitees += 1
                    break  # Sortir de la boucle des feuilles car on a trouvé et traité la facture
        
        # Logs pour le suivi
        if not feuille_uh_trouvee:
            logger.warning(f"Aucune feuille correspondant à l'UH {uh} n'a été trouvée")
        elif not facture_traitee:
            logger.warning(f"Numéro de facture {facture_num} non trouvé dans la feuille {uh}")
    
    return factures_traitees

# Fonction optimisée pour la recherche des factures avec openpyxl
def optimized_openpyxl_search(self, workbook, validated_rows):
    """Version optimisée de la recherche de factures avec openpyxl"""
    # Compteurs pour le rapport
    factures_traitees = 0
    
    # Pour chaque ligne validée, chercher la feuille correspondant à l'UH
    for row_data in validated_rows:
        uh = row_data['uh']
        facture_num = row_data['facture_num'].strip()
        code_client = row_data['code_client']
        code_chorus = row_data['code_chorus']
        
        # Extraire le numéro pur de la facture (sans préfixe)
        facture_num_pur = None
        if "facture n°" in facture_num.lower():
            facture_num_pur = facture_num.lower().replace("facture n°", "").strip()
        elif "facture n" in facture_num.lower():
            facture_num_pur = facture_num.lower().replace("facture n", "").strip()
        elif "facture" in facture_num.lower():
            facture_num_pur = facture_num.lower().replace("facture", "").strip()
        else:
            facture_num_pur = facture_num.lower().strip()
        
        # Log pour débogage
        logger.info(f"Traitement de la facture {facture_num} (UH: {uh}, numéro pur: {facture_num_pur})")
        
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
                
                # Chercher le numéro de facture exact dans cette feuille spécifique
                facture_trouvee = False
                facture_row = None
                facture_col = None
                
                # Rechercher uniquement dans les colonnes 8 (H) et 16 (P) où se trouvent les numéros de facture
                colonnes_facture = [8, 16]  # Colonnes H et P
                lignes_vides_consecutives = 0
                max_lignes_vides = 20  # Arrêter après 20 lignes vides consécutives
                
                for row in range(1, min(2000, sheet.max_row + 1)):  # Limite maximale de 2000 lignes
                    ligne_contient_donnees = False
                    
                    for col in colonnes_facture:  # Uniquement les colonnes où se trouvent les factures
                        try:
                            cell_value = sheet.cell(row=row, column=col).value
                            if cell_value is not None:  # Vérifier que la cellule n'est pas vide
                                ligne_contient_donnees = True
                                # Convertir en string, quelle que soit la valeur
                                cell_text = str(cell_value).lower().strip()
                                
                                # Extraire le numéro pur de la cellule et le normaliser
                                cell_num_pur = None
                                if "facture n°" in cell_text:
                                    cell_num_pur = cell_text.replace("facture n°", "").strip()
                                elif "facture n" in cell_text:
                                    cell_num_pur = cell_text.replace("facture n", "").strip()
                                elif "facture" in cell_text:
                                    cell_num_pur = cell_text.replace("facture", "").strip()
                                else:
                                    cell_num_pur = cell_text
                                
                                # Normaliser les numéros en supprimant les espaces et caractères non numériques
                                normalized_cell_num = ''.join(c for c in cell_num_pur if c.isdigit())
                                normalized_facture_num = ''.join(c for c in facture_num_pur if c.isdigit())
                                
                                logger.debug(f"Comparaison: '{normalized_cell_num}' vs '{normalized_facture_num}'" +
                                            f" (original: '{cell_text}' vs '{facture_num}')")
                                
                                # Vérification avec plusieurs méthodes
                                if (cell_num_pur == facture_num_pur or  # Méthode 1: Comparaison exacte
                                    normalized_cell_num == normalized_facture_num):  # Méthode 2: Comparaison des chiffres uniquement
                                    facture_trouvee = True
                                    facture_row = row
                                    facture_col = col
                                    logger.info(f"Numéro de facture {facture_num} trouvé dans la cellule ({row}, {col}): '{cell_text}'")
                                    break
                        except Exception as e:
                            logger.debug(f"Erreur lors de la lecture de la cellule ({row}, {col}): {str(e)}")
                            continue
                    
                    # Si on a trouvé la facture, on peut sortir de la boucle des lignes
                    if facture_trouvee:
                        break
                    
                    # Vérifier si la ligne contient des données pour gérer le compteur de lignes vides
                    if not ligne_contient_donnees:
                        lignes_vides_consecutives += 1
                    else:
                        lignes_vides_consecutives = 0  # Réinitialiser le compteur si on trouve des données
                    
                    # Arrêter la recherche après 20 lignes vides consécutives
                    if lignes_vides_consecutives >= max_lignes_vides:
                        logger.info(f"Arrêt de la recherche après {max_lignes_vides} lignes vides consécutives à la ligne {row}")
                        break
                
                # Si on a trouvé le numéro de facture exact, insérer les codes
                if facture_trouvee:
                    # Insérer le code client: 6 cellules à gauche et 2 cellules en dessous du numéro de facture
                    if code_client:
                        client_cell_row = facture_row + 2
                        client_cell_col = facture_col - 6
                        if client_cell_col < 1:  # Vérifier que la colonne est valide
                            client_cell_col = 1
                        
                        # Log avant insertion pour vérification
                        logger.info(f"Position du numéro de facture: ({facture_row}, {facture_col})")
                        logger.info(f"Position calculée pour le code client: ({client_cell_row}, {client_cell_col})")
                        
                        sheet.cell(row=client_cell_row, column=client_cell_col).value = code_client
                        logger.info(f"Code client {code_client} inséré dans la cellule ({client_cell_row}, {client_cell_col})")
                    
                    # Insérer le code chorus: 6 cellules à gauche et 1 cellule en dessous du numéro de facture
                    if code_chorus:
                        chorus_cell_row = facture_row + 1
                        chorus_cell_col = facture_col - 6
                        if chorus_cell_col < 1:  # Vérifier que la colonne est valide
                            chorus_cell_col = 1
                        
                        # Log avant insertion pour vérification
                        logger.info(f"Position calculée pour le code chorus: ({chorus_cell_row}, {chorus_cell_col})")
                        
                        sheet.cell(row=chorus_cell_row, column=chorus_cell_col).value = code_chorus
                        logger.info(f"Code chorus {code_chorus} inséré dans la cellule ({chorus_cell_row}, {chorus_cell_col})")
                    
                    facture_traitee = True
                    factures_traitees += 1
                    break  # Sortir de la boucle des feuilles car on a trouvé et traité la facture
        
        # Logs pour le suivi
        if not feuille_uh_trouvee:
            logger.warning(f"Aucune feuille correspondant à l'UH {uh} n'a été trouvée")
        elif not facture_traitee:
            logger.warning(f"Numéro de facture {facture_num} non trouvé dans la feuille {uh}")
    
    return factures_traitees
