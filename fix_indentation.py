#!/usr/bin/env python3
"""
Script pour corriger l'indentation problématique dans setup_database_interface
"""

def fix_indentation():
    """Corrige l'indentation dans setup_database_interface"""
    
    # Lire le fichier main.py
    with open('main.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Trouver et corriger les lignes problématiques autour de la ligne 1379
    for i, line in enumerate(lines):
        # Corriger l'indentation des lignes ajoutées incorrectement
        if 'self.search_interface_ready = True' in line and not line.startswith('            '):
            lines[i] = '            self.search_interface_ready = True\n'
            print(f"Correction indentation ligne {i+1}: search_interface_ready")
        
        elif 'logger.info("Interface de recherche initialisée et prête")' in line and not line.startswith('            '):
            lines[i] = '            logger.info("Interface de recherche initialisée et prête")\n'
            print(f"Correction indentation ligne {i+1}: logger.info")
    
    # Écrire le fichier corrigé
    with open('main.py', 'w', encoding='utf-8') as f:
        f.writelines(lines)
    
    print("Correction de l'indentation terminée !")

if __name__ == '__main__':
    fix_indentation()
