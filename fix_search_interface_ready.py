#!/usr/bin/env python3
"""
Script pour supprimer la ligne restante qui référence search_interface_ready
"""

def fix_search_interface_ready():
    """Supprime la ligne problématique restante"""
    
    # Lire le fichier main.py
    with open('main.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    cleaned_lines = []
    
    for i, line in enumerate(lines):
        # Supprimer la ligne qui référence search_interface_ready
        if 'if not self.search_interface_ready and not force:' in line:
            print(f"Suppression ligne {i+1}: {line.strip()}")
            continue
        elif 'return' in line and i > 0 and 'search_interface_ready' in lines[i-1]:
            print(f"Suppression ligne {i+1}: {line.strip()} (return associé)")
            continue
        else:
            cleaned_lines.append(line)
    
    # Écrire le fichier corrigé
    with open('main.py', 'w', encoding='utf-8') as f:
        f.writelines(cleaned_lines)
    
    print("Correction terminée !")
    print("La référence problématique à search_interface_ready a été supprimée.")

if __name__ == '__main__':
    fix_search_interface_ready()
