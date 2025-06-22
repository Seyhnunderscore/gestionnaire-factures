#!/usr/bin/env python3
"""
Script pour supprimer l'ancienne version de filter_database qui utilise search_input
et corrige les messages d'erreur répétitifs.
"""

import re

def fix_duplicate_filter_database():
    """Supprime l'ancienne version de filter_database problématique"""
    
    # Lire le fichier main.py
    with open('main.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Supprimer la ligne qui connecte search_input à filter_database (ligne ~4663)
    content = re.sub(
        r'\s*self\.search_input\.textChanged\.connect\(self\.filter_database\)\s*\n',
        '',
        content
    )
    
    # Supprimer l'ancienne définition de filter_database (lignes ~4897-4964)
    # Cette version utilise search_input au lieu de db_search_edit
    old_filter_pattern = r'''    def filter_database\(self\):\s*
        """Filtre la base de données en fonction de la recherche"""\s*
        try:\s*
            # Vérifier si search_input est initialisé\s*
            if not hasattr\(self, 'search_input'\) or self\.search_input is None:\s*
                logger\.debug\("Le champ de recherche n'est pas encore initialisé"\)\s*
                return\s*
                \s*
            search_text = self\.search_input\.text\(\)\.strip\(\)\.lower\(\)\s*
            logger\.debug\(f"Recherche dans la base de données: '\{search_text\}'"\)\s*
            \s*
            # Vérifier si le tableau est initialisé\s*
            if not hasattr\(self, 'db_table'\) or self\.db_table is None:\s*
                logger\.error\("Le tableau de la base de données n'est pas initialisé"\)\s*
                return\s*
            \s*
            # Vérifier si la base de données est chargée\s*
            if not hasattr\(self, 'database'\) or not self\.database\._loaded:\s*
                logger\.info\("Chargement de la base de données\.\.\."\)\s*
                self\.database\.ensure_loaded\(lambda: self\.filter_database\(\)\)\s*
                return\s*
            \s*
            # Désactiver les mises à jour de l'interface pour améliorer les performances\s*
            self\.db_table\.setUpdatesEnabled\(False\)\s*
            \s*
            # Si le champ de recherche est vide, afficher toutes les lignes\s*
            if not search_text:\s*
                logger\.debug\("Champ de recherche vide, affichage de toutes les lignes"\)\s*
                for row in range\(self\.db_table\.rowCount\(\)\):\s*
                    self\.db_table\.setRowHidden\(row, False\)\s*
                return\s*
            \s*
            # Nombre de correspondances trouvées\s*
            matches = 0\s*
            \s*
            # Parcourir toutes les lignes du tableau\s*
            for row in range\(self\.db_table\.rowCount\(\)\):\s*
                hide_row = True\s*
                \s*
                # Vérifier chaque colonne de la ligne\s*
                for col in range\(self\.db_table\.columnCount\(\)\):\s*
                    item = self\.db_table\.item\(row, col\)\s*
                    if item is not None and search_text in item\.text\(\)\.lower\(\):\s*
                        hide_row = False\s*
                        matches \+= 1\s*
                        break\s*
                \s*
                # Masquer ou afficher la ligne selon les résultats\s*
                self\.db_table\.setRowHidden\(row, hide_row\)\s*
            \s*
            logger\.debug\(f"\{matches\} correspondance\(s\) trouvée\(s\) pour '\{search_text\}'"\)\s*
            \s*
        except Exception as e:\s*
            logger\.error\(f"Erreur lors du filtrage de la base de données: \{e\}", exc_info=True\)\s*
            # En cas d'erreur, afficher toutes les lignes\s*
            if hasattr\(self, 'db_table'\):\s*
                for row in range\(self\.db_table\.rowCount\(\)\):\s*
                    self\.db_table\.setRowHidden\(row, False\)\s*
                \s*
        finally:\s*
            # Réactiver les mises à jour de l'interface\s*
            if hasattr\(self, 'db_table'\):\s*
                self\.db_table\.setUpdatesEnabled\(True\)\s*
                self\.db_table\.viewport\(\)\.update\(\)\s*
                \s*
            # Mettre à jour la barre d'état\s*
            if hasattr\(self, 'statusBar'\):\s*
                self\.statusBar\(\)\.showMessage\(f"Recherche terminée - \{matches if 'matches' in locals\(\) else '\?'\} résultat\(s\)", 3000\)'''
    
    # Approche plus simple : chercher le bloc complet à supprimer
    lines = content.split('\n')
    new_lines = []
    i = 0
    skip_until_next_def = False
    
    while i < len(lines):
        line = lines[i]
        
        # Détecter le début de l'ancienne version de filter_database
        if 'def filter_database(self):' in line and 'search_input' in ''.join(lines[i:i+10]):
            # Trouver la fin de cette méthode (prochaine méthode ou fin de classe)
            skip_until_next_def = True
            print(f"Suppression de l'ancienne version filter_database à la ligne {i+1}")
        elif skip_until_next_def and (line.strip().startswith('def ') or 
                                     line.strip().startswith('class ') or
                                     (line.strip() and not line.startswith('    '))):
            # Fin de la méthode à supprimer
            skip_until_next_def = False
            new_lines.append(line)
        elif not skip_until_next_def:
            # Supprimer aussi la ligne qui connecte search_input
            if 'self.search_input.textChanged.connect(self.filter_database)' not in line:
                new_lines.append(line)
            else:
                print(f"Suppression de la connexion search_input à la ligne {i+1}")
        
        i += 1
    
    # Écrire le fichier corrigé
    with open('main.py', 'w', encoding='utf-8') as f:
        f.write('\n'.join(new_lines))
    
    print("Correction terminée !")
    print("L'ancienne version de filter_database utilisant search_input a été supprimée.")
    print("Les messages d'erreur répétitifs devraient maintenant disparaître.")

if __name__ == '__main__':
    fix_duplicate_filter_database()
