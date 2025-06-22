from PyQt5.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QApplication
from PyQt5.QtCore import Qt, QPoint, QEvent
from PyQt5.QtGui import QIcon
import logging

logger = logging.getLogger(__name__)

class FloatingWindow(QMainWindow):
    def __init__(self, dock_widget):
        super().__init__(None)  # Pas de parent !
        self.dock_widget = dock_widget
        self.setWindowTitle(dock_widget.windowTitle())
        self.setObjectName(f"{dock_widget.objectName()}_floating")
        
        logger.info(f"Création d'une fenêtre flottante pour {dock_widget.windowTitle()}")
        
        # Initialiser les variables pour le drag & drop
        self.dragging = False
        self.mouse_pos = None
        
        # Créer un conteneur pour le widget
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Extraire le widget du dock mais sauvegarder une référence
        self.original_widget = dock_widget.widget()
        if self.original_widget:
            # Détacher le widget du dock
            self.original_widget.setParent(None)
            # L'ajouter à notre conteneur
            layout.addWidget(self.original_widget)
            logger.info(f"Widget ajouté à la fenêtre flottante: {self.original_widget}")
        else:
            logger.warning("Pas de widget dans le dock")
            
        self.setCentralWidget(container)
        
        # Ensuite, configurer les flags de fenêtre
        # Appliquer les flags AVANT tout show
        flags = Qt.WindowFlags()
        flags |= Qt.Window
        flags |= Qt.WindowMinimizeButtonHint
        flags |= Qt.WindowMaximizeButtonHint
        flags |= Qt.WindowCloseButtonHint
        flags |= Qt.WindowSystemMenuHint
        flags |= Qt.WindowTitleHint
        
        # Appliquer les flags
        self.setWindowFlags(flags)
        
        # Configurer les attributs de fenêtre
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setWindowModality(Qt.NonModal)

    def closeEvent(self, event):
        # Gérer la fermeture de la fenêtre flottante et restaurer le widget dans le dock
        try:
            logger.info("Fermeture de la fenêtre flottante")
            
            # Récupérer le widget de notre conteneur
            container = self.centralWidget()
            if container and container.layout().count() > 0:
                widget = container.layout().itemAt(0).widget()
                if widget:
                    logger.info(f"Récupération du widget: {widget}")
                    # Détacher le widget de notre conteneur
                    widget.setParent(None)
                    # Le réattacher au dock
                    self.dock_widget.setWidget(widget)
                    logger.info("Widget réattaché au dock")
            
            # Réafficher le dock d'origine
            self.dock_widget.setFloating(False)
            self.dock_widget.show()
            
            # Accepter l'événement de fermeture
            event.accept()
            
        except Exception as e:
            logger.error(f"Erreur lors de la fermeture de la fenêtre flottante: {e}")
            event.accept()  # Toujours accepter l'événement pour éviter un blocage
            
    def mousePressEvent(self, event):
        # Gérer le début du drag & drop
        if event.button() == Qt.LeftButton:
            self.dragging = True
            self.mouse_pos = event.globalPos()
        super().mousePressEvent(event)
    
    def mouseMoveEvent(self, event):
        # Gérer le déplacement pendant le drag & drop
        if self.dragging and self.mouse_pos is not None:
            try:
                # Calculer le déplacement
                delta = event.globalPos() - self.mouse_pos
                # Déplacer la fenêtre
                self.move(self.pos() + delta)
                # Mettre à jour la position de la souris
                self.mouse_pos = event.globalPos()
            except Exception as e:
                print(f"Erreur lors du déplacement de la fenêtre: {e}")
                # En cas d'erreur, arrêter le drag & drop
                self.dragging = False
                self.mouse_pos = None
        super().mouseMoveEvent(event)
    
    def mouseReleaseEvent(self, event):
        # Gérer la fin du drag & drop
        if event.button() == Qt.LeftButton and self.dragging:
            self.dragging = False
            self.mouse_pos = None
        super().mouseReleaseEvent(event)
