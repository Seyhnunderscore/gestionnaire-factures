def update_database_from_table(self):
    """Met à jour la base de données à partir du tableau"""
    try:
        data = []
        for row in range(self.db_table.rowCount()):
            item_data = {
                'nom': self.db_table.item(row, 0).text() if self.db_table.item(row, 0) else "",
                'code_client': self.db_table.item(row, 1).text() if self.db_table.item(row, 1) else "",
                'code_chorus': self.db_table.item(row, 2).text() if self.db_table.item(row, 2) else "",
                'adresse': self.db_table.item(row, 3).text() if self.db_table.item(row, 3) else ""
            }
            data.append(item_data)
        
        # Mettre à jour la base de données
        self.database.data = data
        self.database.save()
        logger.info("Base de données mise à jour avec succès depuis le tableau")
        
    except Exception as e:
        logger.error(f"Erreur lors de la mise à jour de la base de données: {e}")
        QMessageBox.critical(self, "Erreur", f"Erreur lors de l'exportation de la base de données : {str(e)}")
