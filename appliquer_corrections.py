#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Script pour appliquer les corrections au fichier main.py afin d'éviter l'ouverture automatique
des fichiers Excel au démarrage de l'application.
"""

import os
import re
import sys
import shutil
from datetime import datetime

def backup_file(file_path):
    """Crée une sauvegarde du fichier avant modification"""
    backup_path = f"{file_path}.bak.{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    shutil.copy2(file_path, backup_path)
    print(f"Sauvegarde créée: {backup_path}")
    return backup_path

def find_init_method(content):
    """Trouve la méthode __init__ dans le contenu du fichier"""
    init_pattern = r'def\s+__init__\s*\(\s*self\s*(?:,\s*[^)]*\s*)?\):'
    match = re.search(init_pattern, content)
    if match:
        return match.start()
    return -1

def find_load_state_async_call(content, init_start):
    """Trouve l'appel à load_state_async dans la méthode __init__"""
    # Trouver la fin de la méthode __init__
    next_def = content.find("\ndef ", init_start + 1)
    if next_def == -1:
        next_def = len(content)
    
    # Chercher l'appel à QTimer.singleShot dans la méthode __init__
    init_content = content[init_start:next_def]
    timer_pattern = r'QTimer\.singleShot\s*\(\s*\d+\s*,\s*self\.load_state_async\s*\)'
    match = re.search(timer_pattern, init_content)
    if match:
        # Calculer la position absolue dans le contenu complet
        return init_start + match.start(), init_start + match.end()
    return -1, -1

def find_load_state_async_method(content):
    """Trouve la méthode load_state_async dans le contenu du fichier"""
    method_pattern = r'def\s+load_state_async\s*\(\s*self\s*\):'
    match = re.search(method_pattern, content)
    if match:
        return match.start()
    return -1

def get_corrected_load_state_async():
    """Retourne le code corrigé pour la méthode load_state_async"""
    return """def load_state_async(self):
    \"\"\"Charge l'état de l'application de manière asynchrone\"\"\"
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
        logger.error(f"Erreur lors du chargement asynchrone de l'état: {{str(e)}}")
"""

def apply_corrections(file_path):
    """Applique les corrections au fichier main.py"""
    # Créer une sauvegarde du fichier
    backup_path = backup_file(file_path)
    
    # Lire le contenu du fichier
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Trouver la méthode __init__
    init_pos = find_init_method(content)
    if init_pos == -1:
        print("Erreur: Méthode __init__ non trouvée dans le fichier.")
        return False
    
    # Trouver l'appel à load_state_async dans la méthode __init__
    call_start, call_end = find_load_state_async_call(content, init_pos)
    if call_start == -1:
        print("Avertissement: Appel à load_state_async non trouvé dans la méthode __init__.")
    else:
        # Commenter l'appel à load_state_async
        content = content[:call_start] + "# " + content[call_start:call_end] + " # Désactivé pour éviter l'ouverture automatique" + content[call_end:]
        print("Appel à load_state_async commenté dans la méthode __init__.")
    
    # Trouver la méthode load_state_async
    method_pos = find_load_state_async_method(content)
    if method_pos == -1:
        print("Erreur: Méthode load_state_async non trouvée dans le fichier.")
        # Si la méthode n'existe pas, on peut l'ajouter à la fin du fichier
        content += "\n\n" + get_corrected_load_state_async()
        print("Méthode load_state_async ajoutée à la fin du fichier.")
    else:
        # Trouver la fin de la méthode load_state_async
        next_def = content.find("\ndef ", method_pos + 1)
        if next_def == -1:
            next_def = len(content)
        
        # Remplacer la méthode load_state_async par la version corrigée
        content = content[:method_pos] + get_corrected_load_state_async() + content[next_def:]
        print("Méthode load_state_async remplacée par la version corrigée.")
    
    # Écrire le contenu modifié dans le fichier
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"Corrections appliquées avec succès au fichier: {file_path}")
    return True

def main():
    """Fonction principale"""
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        # Utiliser le fichier main.py dans le répertoire courant par défaut
        file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py")
    
    if not os.path.exists(file_path):
        print(f"Erreur: Le fichier {file_path} n'existe pas.")
        return 1
    
    if apply_corrections(file_path):
        print("Toutes les corrections ont été appliquées avec succès.")
        return 0
    else:
        print("Des erreurs se sont produites lors de l'application des corrections.")
        return 1

if __name__ == "__main__":
    sys.exit(main())
