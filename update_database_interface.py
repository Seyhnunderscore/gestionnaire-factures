import re

# Définir les modifications à apporter
new_code = '''    def setup_database_interface(self):
        """Configure l'interface de la base de données"""
        # Section de recherche avancée
        search_group = QGroupBox("Recherche")
        search_group.setStyleSheet(f"QGroupBox {{ border: 1px solid {CYBERPUNK_COLORS['border']}; border-radius: 5px; margin-top: 10px; }} QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 5px; }}")
        search_layout = QVBoxLayout(search_group)
        
        # Barre de recherche principale
        main_search_layout = QHBoxLayout()
        self.db_search_edit = QLineEdit()
        self.db_search_edit.setPlaceholderText("Rechercher dans la base de données...")
        self.db_search_edit.textChanged.connect(self.filter_database)
        self.db_search_edit.setClearButtonEnabled(True)  # Ajoute un bouton pour effacer le texte
        main_search_layout.addWidget(self.db_search_edit)
        
        # Bouton d'ajout
        add_btn = QPushButton("+")
        add_btn.setFixedSize(30, 30)
        add_btn.setToolTip("Ajouter une nouvelle entrée")
        add_btn.clicked.connect(self.add_database_entry)
        main_search_layout.addWidget(add_btn)
        
        search_layout.addLayout(main_search_layout)
        
        # Options de filtrage avancées
        filter_layout = QHBoxLayout()
        
        # Sélecteur de catégorie
        self.filter_category = QComboBox()
        self.filter_category.addItem("Tous les champs")
        self.filter_category.addItem("Nom")
        self.filter_category.addItem("Code client")
        self.filter_category.addItem("Code chorus")
        self.filter_category.addItem("Adresse")
        self.filter_category.currentIndexChanged.connect(self.filter_database)
        filter_layout.addWidget(QLabel("Filtrer par:"))
        filter_layout.addWidget(self.filter_category)
        
        # Option de correspondance exacte
        self.exact_match = QCheckBox("Correspondance exacte")
        self.exact_match.stateChanged.connect(self.filter_database)
        filter_layout.addWidget(self.exact_match)
        
        # Ajouter un spacer pour aligner à gauche
        filter_layout.addStretch()
        
        search_layout.addLayout(filter_layout)
        
        # Compteur de résultats
        self.results_count = QLabel("0 résultats")
        self.results_count.setAlignment(Qt.AlignmentFlag.AlignRight)
        search_layout.addWidget(self.results_count)
        
        self.database_layout.addWidget(search_group)'''

# Lire le fichier
with open('main.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Chercher le motif à remplacer
pattern = r'    def setup_database_interface\(self\):\s+"""Configure l\'interface de la base de données"""\s+# Barre de recherche.*?self\.database_layout\.addLayout\(search_layout\)'
replacement = new_code

# Utiliser une expression régulière pour remplacer le contenu
modified_content = re.sub(pattern, replacement, content, flags=re.DOTALL)

# Écrire le contenu modifié dans le fichier
with open('main.py', 'w', encoding='utf-8') as f:
    f.write(modified_content)

print('Modifications appliquées avec succès dans main.py')
