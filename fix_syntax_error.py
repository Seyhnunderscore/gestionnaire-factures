#!/usr/bin/env python3
"""
Script pour corriger l'erreur de syntaxe causée par la suppression mal formatée
et remplacer search_input par db_search_edit dans les parties restantes du code.
"""

def fix_syntax_and_search_input():
    """Corrige la syntaxe et remplace search_input par db_search_edit"""
    
    # Lire le fichier main.py
    with open('main.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    new_lines = []
    
    for i, line in enumerate(lines):
        # Corriger la ligne 4662 qui a été mal formatée
        if 'self.search_input.setPlaceholderText("Rechercher dans la base de données...")        search_layout.addWidget(search_label)' in line:
            # Séparer en deux lignes correctes
            new_lines.append('        self.db_search_edit.setPlaceholderText("Rechercher dans la base de données...")')
            new_lines.append('        search_layout.addWidget(search_label)')
            print(f"Correction de la ligne {i+1}: séparation de la ligne mal formatée")
        
        # Remplacer toutes les autres occurrences de search_input par db_search_edit
        elif 'search_input' in line:
            corrected_line = line.replace('search_input', 'db_search_edit')
            new_lines.append(corrected_line)
            print(f"Ligne {i+1}: search_input remplacé par db_search_edit")
        else:
            new_lines.append(line)
    
    # Écrire le fichier corrigé
    with open('main.py', 'w', encoding='utf-8') as f:
        f.write('\n'.join(new_lines))
    
    print("Correction de la syntaxe terminée !")
    print("Toutes les références à search_input ont été remplacées par db_search_edit.")

if __name__ == '__main__':
    fix_syntax_and_search_input()
