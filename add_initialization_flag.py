#!/usr/bin/env python3
"""
Script pour ajouter un flag d'initialisation plus robuste à l'interface de recherche
pour éviter complètement les appels prématurés à filter_database.
"""

import re

def add_initialization_flag():
    """Ajoute un flag d'initialisation robuste"""
    
    # Lire le fichier main.py
    with open('main.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Ajouter le flag d'initialisation dans __init__
    init_pattern = r'(def __init__\(self.*?\n.*?super\(\).__init__\(\).*?\n)'
    init_replacement = r'\1        # Flag d\'initialisation de l\'interface de recherche\n        self.search_interface_ready = False\n'
    
    if 'self.search_interface_ready = False' not in content:
        content = re.sub(init_pattern, init_replacement, content, flags=re.DOTALL)
        print("Ajout du flag d'initialisation dans __init__")
    
    # Modifier setup_database_interface pour marquer l'interface comme prête
    if 'self.search_interface_ready = True' not in content:
        setup_db_pattern = r'(def setup_database_interface\(self\):.*?self\.connect_search_signals\(\))'
        setup_db_replacement = r'\1\n        # Marquer l\'interface de recherche comme prête\n        self.search_interface_ready = True\n        logger.info("Interface de recherche initialisée et prête")'
        
        content = re.sub(setup_db_pattern, setup_db_replacement, content, flags=re.DOTALL)
        print("Ajout du marquage d'interface prête dans setup_database_interface")
    
    # Modifier filter_database pour utiliser le flag
    filter_pattern = r'(def filter_database\(self, force=False\):.*?""".*?""")'
    filter_replacement = r'''\1
        # Vérifier si l'interface de recherche est complètement initialisée
        if not self.search_interface_ready and not force:
            logger.debug("Interface de recherche pas encore complètement initialisée")
            return'''
    
    if 'self.search_interface_ready' not in content.split('def filter_database')[1].split('def ')[0]:
        content = re.sub(filter_pattern, filter_replacement, content, flags=re.DOTALL)
        print("Ajout de la vérification du flag dans filter_database")
    
    # Écrire le fichier modifié
    with open('main.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print("Amélioration de la robustesse terminée !")
    print("Un flag d'initialisation a été ajouté pour éviter les appels prématurés.")

if __name__ == '__main__':
    add_initialization_flag()
