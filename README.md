# Gestionnaire de Factures

Ce programme permet d'automatiser le remplissage des codes clients et codes chorus dans les factures à partir d'une base de données Excel.

## Fonctionnalités

- Import de bases de données Excel
- Gestion automatique des correspondances entre factures et base de données
- Interface utilisateur intuitive
- Sauvegarde automatique des données
- Gestion des mises à jour de la base de données
- Statistiques sur les correspondances trouvées
- Enrichissement automatique de la base de données
- Mode de test sans connexion à Outlook

## Installation

### Option 1 : Exécutable Windows

1. Exécutez le fichier `build_exe.bat` pour créer l'exécutable
2. L'exécutable se trouvera dans le dossier `dist\Gestionnaire de Factures`
3. Double-cliquez sur `Gestionnaire de Factures.exe` pour lancer l'application

### Option 2 : Exécution depuis le code source

1. Assurez-vous d'avoir Python 3.8 ou supérieur installé
2. Installez les dépendances :
```bash
pip install -r requirements.txt
```

3. Lancez le programme :
```bash
python main.py
```

## Utilisation

1. Pour ajouter une base de données :
   - Cliquez sur "Charger une base de données Excel"
   - Sélectionnez votre fichier Excel
   - Associez les colonnes aux types de données correspondants

2. Pour traiter des factures :
   - Cliquez sur "Charger un fichier de facturation"
   - Sélectionnez votre fichier de facturation
   - Le programme traitera automatiquement les factures et affichera les correspondances

## Structure des fichiers

- Les factures doivent contenir le mot "Intitulé" dans une cellule pour être traitées
- La base de données est sauvegardée dans un fichier `database.json`
- Les fichiers traités sont sauvegardés avec le suffixe "_updated"
