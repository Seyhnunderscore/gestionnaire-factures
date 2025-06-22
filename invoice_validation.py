import logging
from PyQt5.QtWidgets import QProgressDialog, QMessageBox, QPushButton, QApplication
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor

# Récupération du logger
logger = logging.getLogger()

class InvoiceValidator:
    """
    Classe utilitaire pour la validation des factures avec barre de progression
    """
    
    @staticmethod
    def validate_invoice_row_safe(self, visual_row):
        """
        Valide une ligne de facture avec une barre de progression détaillée
        
        Args:
            visual_row (int): Index visuel de la ligne à valider
        """
        try:
            # Vérifier si la ligne existe
            if visual_row < 0 or visual_row >= self.invoice_table.rowCount():
                logger.warning(f"Tentative de validation d'une ligne invalide: {visual_row}")
                return
            
            # Création et mise à jour d'une QProgressDialog pour indiquer la progression de la validation
            progress_dialog = QProgressDialog("Validation de la facture en cours...", None, 0, 100, self)
            progress_dialog.setWindowTitle("Validation")
            progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            progress_dialog.setMinimumDuration(0)
            progress_dialog.setCancelButton(None)
            progress_dialog.setMinimumWidth(300)
            progress_dialog.show()
            QApplication.processEvents()
            
            # Étape 1: Récupération des données (25%)
            progress_dialog.setValue(25)
            progress_dialog.setLabelText("Récupération des données de la facture...")
            QApplication.processEvents()
            
            # Récupérer les données de la ligne
            uh_item = self.invoice_table.item(visual_row, 0)
            numero_item = self.invoice_table.item(visual_row, 1)
            client_item = self.invoice_table.item(visual_row, 2)
            
            # Extraire les valeurs
            uh = uh_item.text() if uh_item else ""
            numero = numero_item.text() if numero_item else ""
            
            logger.info(f"Validation de la facture ligne {visual_row}: UH={uh}, N°={numero}")
            
            # Étape 2: Mise à jour de la couleur (50%)
            progress_dialog.setValue(50)
            progress_dialog.setLabelText("Application du style visuel...")
            QApplication.processEvents()
            
            # Changer la couleur de la ligne pour indiquer qu'elle est validée
            color = QColor("#0078d7")  # Bleu pour "Validé"
            
            # Appliquer la couleur à toutes les cellules de la ligne
            for col in range(self.invoice_table.columnCount()):
                item = self.invoice_table.item(visual_row, col)
                if item:
                    item.setBackground(color)
            
            # Étape 3: Mise à jour du bouton (75%)
            progress_dialog.setValue(75)
            progress_dialog.setLabelText("Mise à jour du statut...")
            QApplication.processEvents()
            
            # Mettre à jour le bouton
            button = self.invoice_table.cellWidget(visual_row, 8)
            if isinstance(button, QPushButton):
                button.setStyleSheet("background-color: #87CEFA; color: white; font-weight: bold;")
                button.setText("Validé")
            
            # Étape 4: Mise à jour des statistiques (100%)
            progress_dialog.setValue(90)
            progress_dialog.setLabelText("Mise à jour des statistiques...")
            QApplication.processEvents()
            
            # Mettre à jour les statistiques
            if hasattr(self, 'update_statistics'):
                self.update_statistics()
            
            progress_dialog.setValue(100)
            progress_dialog.setLabelText("Validation terminée avec succès!")
            QApplication.processEvents()
            
            # Fermer la boîte de dialogue après un court délai
            QTimer.singleShot(800, progress_dialog.close)
            
            logger.info(f"Facture ligne {visual_row} validée avec succès")
            
        except Exception as e:
            # Fermer la boîte de dialogue de progression si elle existe
            if 'progress_dialog' in locals() and progress_dialog.isVisible():
                progress_dialog.close()
                
            logger.error(f"Erreur lors de la validation de la facture ligne {visual_row}: {str(e)}", exc_info=True)
            QMessageBox.critical(
                self,
                "Erreur",
                f"Une erreur est survenue lors de la validation de la facture:\n{str(e)}"
            )
    
    @staticmethod
    def on_validate_button_clicked(self):
        """
        Méthode appelée quand un bouton Valider est cliqué - retrouve dynamiquement la ligne
        """
        try:
            # Récupérer le bouton qui a été cliqué
            button = self.sender()
            if not button:
                return
                
            # Trouver la ligne correspondante dans le tableau
            for row in range(self.invoice_table.rowCount()):
                if self.invoice_table.cellWidget(row, 8) == button:
                    # Appeler la méthode de validation sécurisée
                    self.validate_invoice_row_safe(row)
                    break
                    
        except Exception as e:
            logger.error(f"Erreur lors de la détection du bouton de validation: {str(e)}", exc_info=True)
            QMessageBox.critical(
                self,
                "Erreur",
                f"Une erreur est survenue lors de la validation:\n{str(e)}"
            )
