#!/usr/bin/env python3
"""
Script pour corriger la fonction de saisie des codes
"""

import os
import shutil
from datetime import datetime

def fix_saisie_codes():
    """Corrige la détection des lignes bleues dans save_invoice_file"""
    
    main_file = "main.py"
    backup_file = f"main.py.backup_fix_saisie_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    if not os.path.exists(main_file):
        print("Erreur: main.py non trouve")
        return False
    
    print(f"Sauvegarde: {backup_file}")
    shutil.copy2(main_file, backup_file)
    
    try:
        with open(main_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Remplacer l'ancienne détection de couleur
        old_detection = 'if item and item.background().color().rgb() == QColor(173, 216, 230).rgb():'
        new_detection = 'if item and item.background().color().name().lower() == "#0078d7":'
        
        if old_detection in content:
            content = content.replace(old_detection, new_detection)
            print("[OK] Detection de couleur bleue corrigee")
        else:
            print("[ATTENTION] Ancienne detection non trouvee")
            return False
        
        # Aussi corriger le message d'erreur pour être plus précis
        old_message = 'QMessageBox.warning(self, "Attention", "Aucune ligne validée (en bleu) n\'a été trouvée.")'
        new_message = 'QMessageBox.warning(self, "Attention", "Aucune ligne validée (en bleu #0078d7) n\'a été trouvée.\\nVérifiez que vous avez bien validé des lignes avec le bouton Valider.")'
        
        if old_message in content:
            content = content.replace(old_message, new_message)
            print("[OK] Message d'erreur ameliore")
        
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("[OK] Fonction saisie des codes corrigee!")
        return True
        
    except Exception as e:
        print(f"Erreur: {e}")
        if os.path.exists(backup_file):
            shutil.copy2(backup_file, main_file)
        return False

if __name__ == "__main__":
    print("=== Correction saisie des codes ===")
    print()
    print("Probleme identifie:")
    print("- La fonction cherche l'ancienne couleur QColor(173, 216, 230)")
    print("- Mais le systeme utilise maintenant #0078d7 pour les validations")
    print()
    
    if fix_saisie_codes():
        print()
        print("=== Succes ===")
        print("La fonction de saisie des codes detectera maintenant")
        print("correctement les lignes validees en bleu (#0078d7).")
        print()
        print("Relancez l'application pour tester!")
    else:
        print("Echec de la correction.")
