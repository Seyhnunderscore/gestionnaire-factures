# Voici le code amélioré pour l'interface de recherche de la base de données

def setup_database_interface(self):
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
    
    self.database_layout.addWidget(search_group)
    
    # Tableau de la base de données
    self.db_table = QTableWidget()
    self.db_table.setColumnCount(5)  # Ajout d'une colonne pour le numéro de ligne
    self.db_table.setHorizontalHeaderLabels(["#", "Nom", "Code Client", "Code Chorus", "Adresse"])
    
    # Configuration des largeurs des colonnes
    header = self.db_table.horizontalHeader()
    header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)  # Colonne # en largeur fixe
    for i in range(1, 5):
        header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)
    header.setStretchLastSection(True)  # Étirer la dernière section (Adresse)
    self.db_table.setColumnWidth(0, 50)  # Largeur fixe pour la colonne #
    
    self.db_table.verticalHeader().setVisible(False)
    self.db_table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked | 
                                 QTableWidget.EditTrigger.EditKeyPressed |
                                 QTableWidget.EditTrigger.SelectedClicked)
    self.db_table.itemChanged.connect(self.on_db_item_changed)
    
    # Initialisation du tableau complet (utilisé pour le filtrage)
    self.full_db_table = QTableWidget()
    self.full_db_table.setColumnCount(5)  # Ajout d'une colonne pour le numéro de ligne
    self.full_db_table.setHorizontalHeaderLabels(["#", "Nom", "Code Client", "Code Chorus", "Adresse"])
    self.full_db_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
    self.full_db_table.verticalHeader().setVisible(False)
    
    # Charger les données dans les tableaux
    self.load_database_into_table()
    
    self.database_layout.addWidget(self.db_table)
    
    # Boutons d'action
    btn_layout = QHBoxLayout()
    
    import_btn = QPushButton("Importer")
    import_btn.clicked.connect(self.import_database)
    btn_layout.addWidget(import_btn)
    
    export_btn = QPushButton("Exporter")
    export_btn.clicked.connect(self.export_database_to_excel)
    btn_layout.addWidget(export_btn)
    
    clear_btn = QPushButton("Vider")
    clear_btn.clicked.connect(self.clear_database)
    btn_layout.addWidget(clear_btn)
    
    self.database_layout.addLayout(btn_layout)
