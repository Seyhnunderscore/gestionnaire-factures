#!/usr/bin/env python3
"""
Script pour nettoyer complètement les erreurs d'indentation
et revenir à un état fonctionnel
"""

def clean_indentation():
    """Nettoie toutes les erreurs d'indentation"""
    
    # Lire le fichier main.py
    with open('main.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    cleaned_lines = []
    skip_problematic_lines = False
    
    for i, line in enumerate(lines):
        line_content = line.strip()
        
        # Supprimer les lignes ajoutées qui causent des problèmes
        if any(problematic in line_content for problematic in [
            "self.search_interface_ready = False",
            "self.search_interface_ready = True", 
            "Interface de recherche pas encore complètement initialisée",
            "Interface de recherche initialisée et prête"
        ]):
            print(f"Suppression ligne problématique {i+1}: {line_content[:50]}")
            continue
            
        # Garder les autres lignes
        cleaned_lines.append(line)
    
    # Écrire le fichier nettoyé
    with open('main.py', 'w', encoding='utf-8') as f:
        f.writelines(cleaned_lines)
    
    print("Nettoyage terminé !")
    print("Toutes les lignes problématiques ont été supprimées.")

if __name__ == '__main__':
    clean_indentation()
