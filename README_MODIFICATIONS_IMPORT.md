# Modifications apportées au Gestionnaire de Factures

## Résumé des modifications

Les modifications suivantes ont été apportées à l'application "Gestionnaire de Factures" :

1. **Désactivation du chargement automatique des fichiers Excel au démarrage**
   - Modification de la méthode `load_state_async` pour éviter l'ouverture automatique de l'explorateur Windows
   - Modification de la méthode `load_state` pour ne pas charger automatiquement le dernier fichier Excel ouvert
   - Les fichiers Excel ne sont plus chargés automatiquement au lancement de l'application
   - L'explorateur Windows ne s'ouvre plus automatiquement au démarrage

2. **Implémentation de l'importation manuelle des fichiers Excel**
   - Ajout de la méthode `process_invoice_file` pour traiter les fichiers Excel de facturation
   - Modification de la méthode `load_invoice_file` pour appeler `process_invoice_file`
   - Ajout de la méthode `update_invoice_table` pour afficher les factures importées dans l'interface

## Fonctionnalités ajoutées

### Importation de fichiers Excel

L'application permet maintenant d'importer des fichiers Excel de facturation via le bouton "Ouvrir" dans l'interface. Le processus d'importation comprend les étapes suivantes :

1. Ouverture d'une boîte de dialogue pour sélectionner un fichier Excel
2. Traitement du fichier Excel pour extraire les informations de facturation
3. Affichage des factures importées dans un tableau dans l'interface

### Traitement des fichiers Excel

La méthode `process_invoice_file` a été implémentée pour traiter les fichiers Excel de facturation. Elle effectue les opérations suivantes :

1. Affichage d'une boîte de dialogue de progression
2. Chargement du fichier Excel avec openpyxl
3. Parcours de chaque feuille (UH) du fichier Excel
4. Recherche des cellules contenant "Intitulé" pour identifier les factures
5. Extraction des informations de facturation (nom, numéro, adresse, etc.)
6. Création d'un dictionnaire pour chaque facture
7. Ajout des factures à la liste des factures de l'application
8. Mise à jour du tableau des factures dans l'interface

### Affichage des factures

La méthode `update_invoice_table` a été implémentée pour afficher les factures importées dans un tableau dans l'interface. Elle effectue les opérations suivantes :

1. Création du tableau des factures s'il n'existe pas
2. Effacement du tableau
3. Ajout des factures au tableau
4. Ajustement des colonnes pour qu'elles s'adaptent au contenu

## Comment utiliser l'application

1. Lancez l'application en exécutant le fichier `main.py`
2. Cliquez sur le bouton "Ouvrir" dans l'interface
3. Sélectionnez un fichier Excel de facturation
4. Les factures seront importées et affichées dans le tableau

## Dépendances

L'application utilise les bibliothèques suivantes :
- PyQt5 pour l'interface graphique
- openpyxl pour la lecture des fichiers Excel
- pandas pour la manipulation des données
- datetime pour la gestion des dates

## Notes techniques

- La méthode `process_invoice_file` est conçue pour extraire les informations de facturation à partir de fichiers Excel spécifiques. Elle recherche les cellules contenant "Intitulé" et extrait les informations à partir de positions relatives à ces cellules.
- La méthode `update_invoice_table` crée un tableau des factures s'il n'existe pas et l'ajoute à l'interface. Elle vérifie d'abord si l'onglet des factures existe, puis ajoute le tableau à cet onglet ou directement à la fenêtre principale.

## Détails des modifications pour empêcher l'ouverture automatique de l'explorateur Windows

Deux méthodes principales ont été modifiées pour empêcher l'ouverture automatique de l'explorateur Windows au démarrage :

### 1. Modification de la méthode `load_state_async`

La méthode `load_state_async` était appelée au démarrage de l'application via un timer :

```python
# Charger l'état précédent de manière asynchrone
QTimer.singleShot(100, self.load_state_async)
```

Cette méthode contenait le code suivant qui chargeait automatiquement le dernier fichier Excel ouvert :

```python
# Vérifier si un fichier de facture récent doit être chargé
if hasattr(self, 'current_invoice_path') and self.current_invoice_path:
    if os.path.exists(self.current_invoice_path):
        # Charger le fichier de facture de manière asynchrone
        QTimer.singleShot(100, lambda: self.load_invoice_file())
```

Ce code a été modifié pour ne plus charger automatiquement le fichier, mais simplement mettre à jour le titre de la fenêtre et activer le bouton pour ouvrir avec Excel :

```python
# Ne pas charger automatiquement le dernier fichier Excel ouvert au démarrage
# Cela empêche l'ouverture automatique de l'explorateur Windows
if hasattr(self, 'current_invoice_path') and self.current_invoice_path:
    if os.path.exists(self.current_invoice_path):
        # Mettre à jour le titre de la fenêtre avec le nom du fichier
        self.setWindowTitle(f"Gestionnaire de Factures - {os.path.basename(self.current_invoice_path)}")
        # Activer le bouton pour ouvrir avec Excel
        if hasattr(self, 'open_in_excel_btn'):
            self.open_in_excel_btn.setEnabled(True)
```

### 2. Modification de la méthode `load_state`

La méthode `load_state` contenait également du code qui chargeait automatiquement le dernier fichier Excel ouvert :

```python
# Restaurer le fichier de facture récent s'il existe
if 'recent_invoice' in state and os.path.exists(state['recent_invoice']):
    self.current_invoice_path = state['recent_invoice']
    self.load_invoice_file()
```

Ce code a été modifié pour ne plus charger automatiquement le fichier, mais simplement stocker le chemin du fichier :

```python
# Restaurer le chemin du fichier de facture récent s'il existe, mais ne pas le charger automatiquement
if 'recent_invoice' in state and os.path.exists(state['recent_invoice']):
    self.current_invoice_path = state['recent_invoice']
    # Ne pas charger automatiquement le fichier
    # self.load_invoice_file()
```

Grâce à ces modifications, l'explorateur Windows ne s'ouvre plus automatiquement au démarrage de l'application, et les fichiers Excel ne sont chargés que lorsque l'utilisateur clique sur le bouton "Ouvrir" dans l'interface.
