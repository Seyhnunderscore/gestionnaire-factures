#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Script de test pour vérifier le bon fonctionnement de l'importation de la base de données.
Ce script permet de tester les fonctionnalités d'importation JSON et Excel sans passer par l'interface graphique.
"""

import os
import sys
import json
import logging
import pandas as pd
from datetime import datetime

# Configuration du logging
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, 'test_import.log')

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger('test_import')

class DatabaseTester:
    """Classe de test pour l'importation de la base de données"""
    
    def __init__(self):
        """Initialisation du testeur"""
        self.data = {}
        self.db_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'database.json')
        logger.info(f"Fichier de base de données: {self.db_file}")
        
        # Charger la base de données existante si elle existe
        if os.path.exists(self.db_file):
            try:
                with open(self.db_file, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
                logger.info(f"Base de données chargée: {len(self.data)} entrées")
            except Exception as e:
                logger.error(f"Erreur lors du chargement de la base de données: {e}")
                self.data = {}
        else:
            logger.warning(f"Le fichier de base de données n'existe pas: {self.db_file}")
    
    def save_database(self):
        """Sauvegarde la base de données dans un fichier JSON"""
        try:
            with open(self.db_file, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=4)
            logger.info(f"Base de données sauvegardée: {len(self.data)} entrées")
            return True
        except Exception as e:
            logger.error(f"Erreur lors de la sauvegarde de la base de données: {e}")
            return False
    
    def add_entry(self, name, client_code=None, chorus_code=None, address=None):
        """Ajoute ou met à jour une entrée dans la base de données"""
        if not name:
            return False
        
        self.data[name] = {
            'client_code': client_code,
            'chorus_code': chorus_code,
            'address': address
        }
        
        logger.info(f"Entrée ajoutée/mise à jour: {name}")
        return True
    
    def load_from_dataframe(self, df, mapping):
        """Charge des données depuis un DataFrame pandas
        
        Args:
            df (pandas.DataFrame): DataFrame contenant les données à importer
            mapping (dict): Dictionnaire de mapping des colonnes {champ_cible: nom_colonne}
                - name: Nom du client (obligatoire)
                - client_code: Code client (optionnel)
                - chorus_code: Code Chorus (optionnel)
                - address: Adresse (optionnel)
                
        Returns:
            int: Nombre d'entrées ajoutées
        """
        logger.info(f"Début du chargement depuis DataFrame: {len(df)} lignes, mapping: {mapping}")
        
        # Vérifier que le mapping contient au moins le nom
        if not mapping.get('name'):
            logger.error("Mapping invalide: le champ 'name' est obligatoire")
            raise ValueError("Le mapping doit contenir au moins le champ 'name'")
        
        # Compteur d'entrées ajoutées
        entries_added = 0
        errors = 0
        
        # Parcourir les lignes du DataFrame
        for index, row in df.iterrows():
            try:
                # Récupérer les valeurs en fonction du mapping
                name = row.get(mapping['name'])
                
                # Vérifier que le nom n'est pas vide
                if pd.isna(name) or not str(name).strip():
                    logger.warning(f"Ligne {index+1}: Nom vide, ignorée")
                    continue
                
                # Convertir en chaîne de caractères et nettoyer
                name = str(name).strip()
                
                # Récupérer les autres valeurs (optionnelles)
                client_code = row.get(mapping.get('client_code')) if mapping.get('client_code') else None
                chorus_code = row.get(mapping.get('chorus_code')) if mapping.get('chorus_code') else None
                address = row.get(mapping.get('address')) if mapping.get('address') else None
                
                # Convertir les valeurs en chaînes de caractères si elles ne sont pas None
                if client_code is not None and not pd.isna(client_code):
                    client_code = str(client_code).strip()
                else:
                    client_code = None
                    
                if chorus_code is not None and not pd.isna(chorus_code):
                    chorus_code = str(chorus_code).strip()
                else:
                    chorus_code = None
                    
                if address is not None and not pd.isna(address):
                    address = str(address).strip()
                else:
                    address = None
                
                # Ajouter l'entrée à la base de données
                self.add_entry(name, client_code, chorus_code, address)
                entries_added += 1
                
                # Log périodique pour les gros fichiers
                if entries_added % 100 == 0:
                    logger.info(f"Progression: {entries_added} entrées traitées")
                    
            except Exception as e:
                errors += 1
                logger.error(f"Erreur lors du traitement de la ligne {index+1}: {str(e)}")
                # Continuer malgré l'erreur
        
        logger.info(f"Chargement terminé: {entries_added} entrées ajoutées, {errors} erreurs")
        return entries_added
    
    def load_file(self, file_path):
        """Charge un fichier JSON dans la base de données
        
        Args:
            file_path (str): Chemin vers le fichier JSON à charger
            
        Returns:
            bool: True si le chargement a réussi, False sinon
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Vérifier si c'est un objet avec une clé 'data' ou directement une liste/dictionnaire
            if isinstance(data, dict) and 'data' in data:
                data = data['data']
            
            # Vérifier que nous avons des données valides
            if not isinstance(data, dict):
                logger.error(f"Format de données JSON invalide: {type(data)}")
                return False
            
            if not data:
                logger.warning("Le fichier JSON ne contient aucune donnée")
                return False
            
            # Mettre à jour la base de données
            self.data = data
            logger.info(f"Données JSON chargées avec succès: {len(self.data)} entrées")
            return True
            
        except Exception as e:
            logger.error(f"Erreur lors du chargement du fichier JSON: {e}")
            return False
    
    def test_json_import(self, file_path):
        """Teste l'importation d'un fichier JSON
        
        Args:
            file_path (str): Chemin vers le fichier JSON à importer
        """
        logger.info(f"Test d'importation JSON: {file_path}")
        
        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            logger.error(f"Le fichier n'existe pas: {file_path}")
            return
        
        # Charger le fichier
        if self.load_file(file_path):
            # Sauvegarder la base de données
            if self.save_database():
                logger.info("Test d'importation JSON réussi")
                # Afficher quelques exemples d'entrées
                sample_keys = list(self.data.keys())[:3] if self.data else []
                if sample_keys:
                    logger.info(f"Exemples d'entrées importées: {', '.join(sample_keys)}")
            else:
                logger.error("Erreur lors de la sauvegarde de la base de données")
        else:
            logger.error("Erreur lors du chargement du fichier JSON")
    
    def test_excel_import(self, file_path, mapping=None):
        """Teste l'importation d'un fichier Excel
        
        Args:
            file_path (str): Chemin vers le fichier Excel à importer
            mapping (dict, optional): Mapping des colonnes. Si None, un mapping automatique sera tenté.
        """
        logger.info(f"Test d'importation Excel: {file_path}")
        
        # Vérifier que le fichier existe
        if not os.path.exists(file_path):
            logger.error(f"Le fichier n'existe pas: {file_path}")
            return
        
        try:
            # Charger le fichier Excel
            df = pd.read_excel(file_path)
            logger.info(f"Fichier Excel lu avec succès: {len(df)} lignes, colonnes: {df.columns.tolist()}")
            
            # Vérifier que le DataFrame n'est pas vide
            if df.empty:
                logger.warning("Le fichier Excel est vide")
                return
            
            # Si aucun mapping n'est fourni, essayer de trouver automatiquement les colonnes pertinentes
            if mapping is None:
                mapping = {}
                for col in df.columns:
                    col_lower = col.lower()
                    if "nom" in col_lower or "name" in col_lower or "client" in col_lower:
                        mapping['name'] = col
                    elif "code" in col_lower and ("client" in col_lower):
                        mapping['client_code'] = col
                    elif "chorus" in col_lower or "code chorus" in col_lower:
                        mapping['chorus_code'] = col
                    elif "adresse" in col_lower or "address" in col_lower:
                        mapping['address'] = col
                
                logger.info(f"Mapping automatique des colonnes: {mapping}")
                
                # Vérifier qu'au moins le nom est mappé
                if not mapping.get('name'):
                    logger.warning("Aucune colonne trouvée pour le nom")
                    # Prendre la première colonne comme nom
                    mapping['name'] = df.columns[0]
                    logger.info(f"Utilisation de la première colonne comme nom: {mapping['name']}")
            
            # Charger les données dans la base de données
            entries_added = self.load_from_dataframe(df, mapping)
            logger.info(f"Données chargées avec succès: {entries_added} entrées ajoutées")
            
            # Sauvegarder la base de données
            if self.save_database():
                logger.info("Test d'importation Excel réussi")
                # Afficher quelques exemples d'entrées
                sample_keys = list(self.data.keys())[:3] if self.data else []
                if sample_keys:
                    logger.info(f"Exemples d'entrées importées: {', '.join(sample_keys)}")
            else:
                logger.error("Erreur lors de la sauvegarde de la base de données")
                
        except Exception as e:
            logger.error(f"Erreur lors de l'importation Excel: {e}")

def create_test_json():
    """Crée un fichier JSON de test avec quelques entrées"""
    test_data = {
        "Client Test 1": {
            "client_code": "CT001",
            "chorus_code": "CH001",
            "address": "1 rue du Test, 75000 Paris"
        },
        "Client Test 2": {
            "client_code": "CT002",
            "chorus_code": "CH002",
            "address": "2 avenue de l'Exemple, 69000 Lyon"
        },
        "Client Test 3": {
            "client_code": "CT003",
            "chorus_code": "CH003",
            "address": "3 boulevard du Modèle, 33000 Bordeaux"
        }
    }
    
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_data.json')
    
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(test_data, f, ensure_ascii=False, indent=4)
        logger.info(f"Fichier JSON de test créé: {file_path}")
        return file_path
    except Exception as e:
        logger.error(f"Erreur lors de la création du fichier JSON de test: {e}")
        return None

def create_test_excel():
    """Crée un fichier Excel de test avec quelques entrées"""
    test_data = {
        "Nom": ["Client Excel 1", "Client Excel 2", "Client Excel 3"],
        "Code Client": ["CE001", "CE002", "CE003"],
        "Code Chorus": ["CHE001", "CHE002", "CHE003"],
        "Adresse": [
            "10 rue du Test Excel, 75000 Paris",
            "20 avenue de l'Exemple Excel, 69000 Lyon",
            "30 boulevard du Modèle Excel, 33000 Bordeaux"
        ]
    }
    
    df = pd.DataFrame(test_data)
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_data.xlsx')
    
    try:
        df.to_excel(file_path, index=False)
        logger.info(f"Fichier Excel de test créé: {file_path}")
        return file_path
    except Exception as e:
        logger.error(f"Erreur lors de la création du fichier Excel de test: {e}")
        return None

def main():
    """Fonction principale"""
    logger.info("Démarrage des tests d'importation de la base de données")
    
    # Créer un testeur
    tester = DatabaseTester()
    
    # Créer des fichiers de test
    json_file = create_test_json()
    excel_file = create_test_excel()
    
    # Tester l'importation JSON
    if json_file:
        tester.test_json_import(json_file)
    
    # Tester l'importation Excel
    if excel_file:
        # Définir un mapping explicite pour le test
        mapping = {
            'name': 'Nom',
            'client_code': 'Code Client',
            'chorus_code': 'Code Chorus',
            'address': 'Adresse'
        }
        tester.test_excel_import(excel_file, mapping)
    
    logger.info("Tests terminés")

if __name__ == "__main__":
    main()
