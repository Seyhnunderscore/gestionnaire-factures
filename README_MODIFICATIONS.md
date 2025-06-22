# Modifications à apporter au fichier main.py

Ce document explique les modifications à apporter au fichier `main.py` pour éviter l'ouverture automatique des fichiers Excel au démarrage de l'application.

## 1. Modification de la méthode `load_state_async`

La méthode `load_state_async` est responsable du chargement asynchrone de l'état de l'application après le démarrage. Elle est appelée via un `QTimer.singleShot` dans le constructeur `__init__` de la classe `MainWindow`.

Pour éviter l'ouverture automatique des fichiers Excel au démarrage, vous devez modifier cette méthode comme suit :

```python
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
```

## 2. Modification du constructeur `__init__` de la classe `MainWindow`

Si l'appel à `load_state_async` est présent dans le constructeur `__init__` de la classe `MainWindow`, vous pouvez le commenter pour éviter complètement l'appel à cette méthode au démarrage :

```python
def __init__(self):
    super().__init__()
    # ...
    
    # Commenter cette ligne pour éviter le chargement automatique au démarrage
    # QTimer.singleShot(100, self.load_state_async)
    
    # ...
```

## 3. Vérification de la méthode `load_invoice_file`

Assurez-vous que la méthode `load_invoice_file` est correctement implémentée pour ouvrir un fichier Excel uniquement lorsque l'utilisateur clique sur le bouton "Charger une facture" :

```python
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
        
        # Si l'utilisateur a annulé, ne rien faire
        if not file_path:
            return
        
        # Enregistrer le chemin du fichier pour une utilisation ultérieure
        self.current_excel_file = file_path
        
        # Traiter le fichier Excel
        self.process_invoice_file(file_path)
        
        # Mettre à jour l'interface
        self.update_preview_table()
        
        # Sauvegarder l'état de l'application
        self.save_state(show_message=False)
        
    except Exception as e:
        logger.error(f"Erreur lors du chargement du fichier de facturation: {str(e)}")
        QMessageBox.critical(
            self,
            "Erreur de chargement",
            f"Une erreur est survenue lors du chargement du fichier: {str(e)}"
        )
```

## 4. Vérification de la méthode `process_invoice_file`

La méthode `process_invoice_file` doit être correctement implémentée pour traiter les fichiers Excel de facturation. Vous pouvez utiliser la version corrigée du fichier `corrections.py` comme référence.

## Conclusion

En appliquant ces modifications, vous vous assurez que les fichiers Excel ne sont plus chargés automatiquement au démarrage de l'application, mais uniquement lorsque l'utilisateur clique sur le bouton "Charger une facture".
