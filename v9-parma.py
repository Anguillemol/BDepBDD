import pandas as pd
import openpyxl as wb
import sys, io, shutil, tempfile, os, time
import datetime
from unidecode import unidecode

from PyQt6.QtCore import Qt, QSize, QSortFilterProxyModel, QAbstractTableModel, QTimer, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QPainter, QColor, QPixmap, QBrush, QFont
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QCheckBox, QLabel, QDialogButtonBox, QProgressBar, QHeaderView, QMessageBox, QHBoxLayout, QWidget, QLineEdit, QGridLayout, QComboBox, QVBoxLayout, QStackedWidget, QScrollArea, QFrame, QTableView, QSpacerItem 
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File 


""" SplashScreen : Classe gérant l'écran de chargement
    
    {__init__}: Fonction d'initialisation de la classe
    {initUI}: Annexe de l'initialisation de la classe (mise en place des Widgets)
"""
class SplashScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet('''
        #LabelTitle {
            font-size: 60px;
            color: #93deed;
        }

        #LabelDesc {
            font-size: 30px;
            color: #c2ced1;
        }

        #LabelLoading {
            font-size: 30px;
            color: #e8e8eb;
        }

        QFrame {
            background-color: #2F4454;
            color: rgb(220, 220, 220);
        }

        QProgressBar {
            background-color: #DA7B93;
            color: rgb(200, 200, 200);
            border-style: none;
            border-radius: 10px;
            text-align: center;
            font-size: 30px;
        }

        QProgressBar::chunk {
            border-radius: 10px;
            background-color: qlineargradient(spread:pad x1:0, x2:1, y1:0.511364, y2:0.523, stop:0 #1C3334, stop:1 #376E6F);
        }
    ''')
        self.setWindowTitle('Spash Screen Example')
        self.setFixedSize(1100, 500)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        self.counter = 0
        self.n = 100 # total instance

        self.initUI()

        self.timer = QTimer()
        #self.timer.timeout.connect(self.loading)
        #self.timer.start(30)

    def initUI(self):
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.frame = QFrame()
        layout.addWidget(self.frame)

        self.labelTitle = QLabel(self.frame)
        self.labelTitle.setObjectName('LabelTitle')

        
        self.labelTitle.resize(self.width() - 10, 160)
        self.labelTitle.move(0, 40) # x, y
        self.labelTitle.setText('Chargement')
        self.labelTitle.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.labelDescription = QLabel(self.frame)
        self.labelDescription.resize(self.width() - 10, 50)
        self.labelDescription.move(0, self.labelTitle.height())
        self.labelDescription.setObjectName('LabelDesc')
        self.labelDescription.setText('<strong>Working on Task #1</strong>')
        self.labelDescription.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.progressBar = QProgressBar(self.frame)
        self.progressBar.resize(self.width() - 200 - 10, 50)
        self.progressBar.move(100, self.labelDescription.y() + 130)
        self.progressBar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progressBar.setFormat('%p%')
        self.progressBar.setTextVisible(True)
        self.progressBar.setRange(0, self.n)
        self.progressBar.setValue(0)

        self.labelLoading = QLabel(self.frame)
        self.labelLoading.resize(self.width() - 10, 50)
        self.labelLoading.move(0, self.progressBar.y() + 70)
        self.labelLoading.setObjectName('LabelLoading')
        self.labelLoading.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.labelLoading.setText('Chargement...')

        self.labelSubDescription = QLabel(self.frame)
        self.labelSubDescription.resize(self.width() - 10, 50)
        self.labelSubDescription.move(0, self.labelDescription.y()+ 50)
        self.labelSubDescription.setObjectName('LabelSubDesc')
        self.labelSubDescription.setText('')
        self.labelSubDescription.setAlignment(Qt.AlignmentFlag.AlignCenter)

""" LogWindow : Classe gérant l'écran d'authentification

    {__init__}: Fonction d'initialisation de la classe (mise en place des widgets et layouts)

    {PushCo}: Fonction d'authentification, ouvre la fenêtre principale si l'authentification est correcte

    {PushCl}: Fonction de nettoyage des champs de saisie
"""
class logWindow(QWidget):  
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Authentification")
        self.setFixedSize(570,200)
        self.w = None 
        self.role = ""

        ########## Lecture du fichier mot de passe ##########
        self.ddbmdp = pd.read_excel(mdp, sheet_name='MDPFINAUX')

        ########## logV1 ##########
        self.logv1 = QWidget()
        
        self.hLayout1 = QHBoxLayout()

        self.logoLabel1 = QLabel("")
        self.logoLabel1.setMaximumSize(QSize(200,140))
        self.logoLabel1.setPixmap(QPixmap(os.path.join(repertexec,"logo.png")))
        self.logoLabel1.setScaledContents(True)
        self.logoLabel1.setStyleSheet("background-color: white")



        self.hLayout1.addWidget(self.logoLabel1)

        font = QFont()
        font.setPointSize(40)
        font.setBold(True)
        font.setWeight(75)

        ##### Labels #####
        self.user = QLabel("Nom d'utilisateur:")
        self.password = QLabel("Mot de passe:")

        ##### LineEdit #####
        self.inputUser = QLineEdit()
        self.inputPassword = QLineEdit()
        self.inputPassword.setObjectName("button1")
        self.inputPassword.setEchoMode(QLineEdit.EchoMode.Password)

        ##### Buttons #####
        self.buttonCo = QPushButton("Connexion")
        self.buttonCo.setObjectName("button1")
        self.buttonCo.setFont(font)
        self.buttonCo.setStyleSheet("font-size:14px;")
        self.buttonCo.clicked.connect(self.PushCo)
        self.buttonCl = QPushButton("Nettoyer")
        self.buttonCl.clicked.connect(self.PushCl)
        self.buttonCl.setFont(font)
        self.buttonCl.setStyleSheet("font-size:14px;")
        self.inputPassword.returnPressed.connect(self.PushCo)
        
        self.gLayout1 = QGridLayout()

        self.gLayout1.addWidget(self.user, 0, 0, 1, 1)
        self.gLayout1.addWidget(self.password, 1, 0, 1, 1)
        self.gLayout1.addWidget(self.inputUser, 0, 1, 1, 2)
        self.gLayout1.addWidget(self.inputPassword, 1, 1, 1, 2)
        self.gLayout1.addWidget(self.buttonCl, 2, 1, 1, 1)
        self.gLayout1.addWidget(self.buttonCo, 2, 2, 1, 1)

        self.gWidget1 = QWidget()
        self.gWidget1.setLayout(self.gLayout1)

        self.hLayout1.addWidget(self.gWidget1, 0, Qt.AlignmentFlag.AlignVCenter)
        
        self.logv1.setLayout(self.hLayout1)


        ########## logV2 (authentification incorrecte) ##########
        self.logv2 = QWidget()
        
        self.hLayout2 = QHBoxLayout()

        self.logoLabel2 = QLabel()
        self.logoLabel2.setMaximumSize(QSize(200,140))
        self.logoLabel2.setPixmap(QPixmap(os.path.join(repertexec,"logo.png")))
        self.logoLabel2.setScaledContents(True)
        self.logoLabel2.setStyleSheet("background-color: white")
        
        self.hLayout2.addWidget(self.logoLabel2)

        ##### Labels #####
        self.user2 = QLabel("Nom d'utilisateur:")
        self.password2 = QLabel("Mot de passe:")

        self.error = QLabel("Échec d'authentification")
        self.error.setStyleSheet("color: red;")

        ##### LineEdit #####
        self.inputUser2 = QLineEdit()
        self.inputPassword2 = QLineEdit()
        self.inputPassword2.setObjectName("button2")
        self.inputPassword2.setEchoMode(QLineEdit.EchoMode.Password)

        ##### Buttons #####
        self.buttonCo2 = QPushButton("Connexion")
        self.buttonCo2.setObjectName("button2")
        self.buttonCo2.clicked.connect(self.PushCo)
        self.buttonCo2.setFont(font)
        self.buttonCl2 = QPushButton("Nettoyer")
        self.buttonCl2.clicked.connect(self.PushCl)
        self.inputPassword2.returnPressed.connect(self.PushCo)
        self.buttonCl2.setFont(font)
        
        self.gLayout2 = QGridLayout()

        self.gLayout2.addWidget(self.user2, 0, 0, 1, 1)
        self.gLayout2.addWidget(self.password2, 1, 0, 1, 1)
        self.gLayout2.addWidget(self.inputUser2, 0, 1, 1, 2)
        self.gLayout2.addWidget(self.inputPassword2, 1, 1, 1, 2)
        self.gLayout2.addWidget(self.buttonCl2, 2, 1, 1, 1)
        self.gLayout2.addWidget(self.buttonCo2, 2, 2, 1, 1)
        self.gLayout2.addWidget(self.error, 3, 1, 1, 2)   

        self.gWidget2 = QWidget()
        self.gWidget2.setLayout(self.gLayout2)

        self.hLayout2.addWidget(self.gWidget2)

        self.logv2.setLayout(self.hLayout2)
      
        ############## StackedWidget ##############
        self.Stack = QStackedWidget (self)
        self.Stack.addWidget(self.logv1)
        self.Stack.addWidget(self.logv2)

        self.Stack.setCurrentIndex(0)

        layoutHome = QHBoxLayout()
        layoutHome.addWidget(self.Stack)

        self.setLayout(layoutHome)

    ########## Fonction de connexion ##########
    def PushCo (self):
        sender = self.sender()
        name = sender.objectName()

        if name=="button1":
            if self.ddbmdp.loc[self.ddbmdp['Username'] == self.inputUser.text()].empty:
                self.Stack.setCurrentIndex(1)
            else:
                nrow = self.ddbmdp[self.ddbmdp['Username'] == self.inputUser.text()].index.values.astype(int)[0]
                role = self.ddbmdp['Role'][nrow]
                denom = self.ddbmdp['Denom'][nrow]
                if self.inputPassword.text() == self.ddbmdp.loc[self.ddbmdp['Username'] == self.inputUser.text()]['Password'][nrow]:
                    ##### Authentification réussie, ouverture de la fenêtre principale #####
                    self.Stack.setCurrentIndex(2)
                    if self.w is None:
                        self.w=main
                        main.user = self.inputUser.text()
                        main.password = self.inputPassword.text()
                        main.role = role
                        main.denom = denom
                        main.demarrage()
                    self.w.show()
                    self.close()
                else:
                    self.Stack.setCurrentIndex(1)
        if name=="button2":
            if not (self.ddbmdp.loc[self.ddbmdp['Username'] == self.inputUser2.text()].empty):
                nrow = self.ddbmdp[self.ddbmdp['Username'] == self.inputUser2.text()].index.values.astype(int)[0]
                role = self.ddbmdp['Role'][nrow]
                denom = self.ddbmdp['Denom'][nrow]
                if self.inputPassword2.text() == self.ddbmdp.loc[self.ddbmdp['Username'] == self.inputUser2.text()]['Password'][nrow]:
                    ##### Authentification réussie, ouverture de la fenêtre principale #####
                    self.Stack.setCurrentIndex(2)
                    if self.w is None:
                        self.w=main
                        main.user = self.inputUser2.text()
                        main.password = self.inputPassword2.text()
                        main.role = role
                        main.denom = denom
                        main.demarrage()
                        #self.w=mainWindow()
                        #self.w.user = self.inputUser2.text()
                        #self.w.password = self.inputPassword2.text()
                        #self.w.role = role
                    self.w.show()
                    self.close()
                else:
                    
                    self.inputUser2.setText("")
                    self.inputPassword2.setText("")
            else:
                self.inputUser2.setText("")
                self.inputPassword2.setText("")

    ########## Fonction de nettoyage des champs de saisie ##########
    def PushCl (self):
        self.inputUser.setText("")
        self.inputUser2.setText("")
        self.inputPassword.setText("")
        self.inputPassword2.setText("")

""" mainWindow : Classe gérant la fenêtre principale

    {__init__}: Fonction d'initialisation de la classe

    {demarrage}: Fonction de démarrage de la classe, execute les fonctions loadExcel et loadGU

    {loadExcel}: Fonction de récupération des données Excel

    {loadGUI}: Fonction de génération de l'interface

    {center}: Fonction de centrage de la fenêtre dans l'écran

    {creerDepot}: Fonction créant la fenêtre "Création de dépôt" (réservé aux utilisateurs admins et superadmins)

    {modifDepot}: Fonction créant la fenêtre "Modification de dépôt" (réservé aux utilisateurs admins et superadmins)

    {supprimerDepot}: Fonction créant la fenêtre "Suppression de dépôt" (réservé aux utilisateurs admins et superadmins)

    {demandeChangement}: Fonction créant la fenêtre "Demande de changement de données" (réservé aux utilisateurs non-admins)

    {traitementRequetesChangement}: Fonction créant la fenêtre "Traitement des demandes de changement de données" (réservé aux utilisateurs admins et superadmins)

    {chargerModif}: Fonction pour actualiser les données de la fenêtre principale

    {reglages}: Fonction d'accès à la fenêtre de régalge (réservé aux SuperAdmins)

    {saveData}: Fonction sauvegardant les données saisies dans l'outil dans Sharepoint (réservé aux utilisateurs admins et superadmins)
"""
##TODO: SAVEDATA
class mainWindow(QWidget):
    sheetRequetes = pd.DataFrame
    lstParam = []
    def __init__(self):
        super().__init__()

        self.setFixedSize(720,440)

        ##### Important variables #####
        self.user = ""
        self.password = ""
        self.role = ""
        self.denom = ""

        ##### Loading Widget #####

        self.loading = QWidget()
        self.layoutLoading = QGridLayout()

        self.loadButton = QPushButton("Démarrer")
        self.loadButton.resize(600,100)
        self.loadButton.clicked.connect(self.demarrage)

        self.layoutLoading.addWidget(self.loadButton, 0, 0, 3, 3, Qt.AlignmentFlag.AlignCenter)
        self.loading.setLayout(self.layoutLoading)

        ##### Creation of the stackedWidget and main layout #####
        self.mainLayout = QVBoxLayout()

        self.Stack = QStackedWidget (self)
        self.Stack.addWidget(self.loading)
        
        self.Stack.setCurrentIndex(1)
        self.mainLayout.addWidget(self.Stack)

        self.setLayout(self.mainLayout)
               
    def demarrage(self):
        print("Username: " + self.user)
        print("Password: " + self.password)
        print("Role: " + self.role)
        print("Denomination: " + self.denom)
        self.setFixedSize(1280,720)
        self.center()
        ##### Lecture base de données #####
        excel = pd.ExcelFile(bdd)
        lst_sheet = excel.sheet_names
        self.d = {}
        for i in range(len(lst_sheet)):
            if lst_sheet[i] != "Accueil" and lst_sheet[i] != "BDD":
                self.d[lst_sheet[i]] = pd.read_excel(bdd, sheet_name=lst_sheet[i])

        ##### Récupération des données #####
        self.loadExcel()

        ##### Generation de l'interface #####
        self.loadGUI()
        
    def loadExcel(self):
        self.sheet = sheet_Globale
        self.sheet_columns = self.sheet.columns.to_list()

        if self.role not in ["Admin","SuperAdmin"]:
            if self.role == "Région":
                ##### Récupération de toutes les lignes d'un directeur régional et affichables en fonction de la liste des paramètres (ici self.lstParam) #####
                self.sheet_tri = self.sheet.loc[self.sheet['Directeur Régional'] == self.denom, self.lstParam]
                self.model = pandasModel(self.sheet_tri)
            elif self.role == "Dépôt":
                ##### Récupération des lignes d'un directeur de dépôt et affichables en fonction de la liste des paramètres (ici self.lstParam) #####
                self.sheet_tri = self.sheet.loc[self.sheet['Directeur dépôt'] == self.denom, self.lstParam]
                self.model = pandasModel(self.sheet_tri)
        else:
            if self.role == "SuperAdmin":
                self.model = pandasEditableModel(self.sheet)
            elif self.role == "Admin":
                ##### Limitation des colonnes affichables en fonction de la liste des paramètres (ici self.lstParam)#####
                self.sheetParam = self.sheet.loc[:, self.lstParam]
                self.model = pandasEditableModel(self.sheetParam)
            print("Model créé Admin en lecture écriture")

    def loadGUI(self):
        ##### Création de l'interface ADMIN et SuperAdmin #####
        if self.role == "Admin" or self.role == "SuperAdmin":
            print("Creation du GUI ADMIN")

            self.adminGUI = QWidget()

            ##### Top Banner #####
            self.titre = QLabel("Base de données magasin - Brico Dépôt")
            font = QFont()
            font.setPointSize(28)
            font.setBold(True)
            font.setWeight(75)
            self.titre.setFont(font)
            self.titre.setFrameShape(QFrame.Shape.Box)
            self.titre.setFrameShadow(QFrame.Shadow.Plain)
            self.titre.setLineWidth(5)
            self.titre.setStyleSheet("background-color: rgb(191, 189, 184)")

            self.logo = QLabel("")
            self.logo.setMaximumSize(QSize(170, 100))
            self.logo.setPixmap(QPixmap(os.path.join(repertexec,"logo.png")))
            self.logo.setScaledContents(True)
            self.logo.setStyleSheet("background-color: white")

            self.infos = QLabel(self.denom + "\n" + self.role)
            font2 = QFont()
            font2.setPointSize(16)
            self.infos.setFont(font2)
            self.infos.setFrameShape(QFrame.Shape.Box)
            self.infos.setFrameShadow(QFrame.Shadow.Plain)
            self.infos.setLineWidth(5)
            self.infos.setStyleSheet("background-color: rgb(191, 189, 184)")

            self.topBanner = QWidget()

            self.topLayout = QHBoxLayout()
            self.topLayout.setContentsMargins(35, 10, -1, -1)
            self.topLayout.setSpacing(60)
            self.topLayout.addWidget(self.logo)
            self.topLayout.addWidget(self.titre)
            self.topLayout.addWidget(self.infos)
            self.topLayout.addStretch(1)
            
            self.topBanner.setLayout(self.topLayout)

            ##### Table and research bar #####
            self.middle = QWidget()
            self.middleLayout = QHBoxLayout()
            self.middleLayout.setSpacing(20)

            self.searchBar = QLineEdit()
            self.searchBar.setMaximumSize(600,30)
            self.searchBar.setPlaceholderText("Champ de recherche")
            self.searchFont = QFont()
            self.searchFont.setPointSize(16)
            self.searchFont.setBold(False)
            self.searchFont.setWeight(20)

            self.tab = QTableView()

            self.proxy_model = QSortFilterProxyModel()
            self.proxy_model.setFilterKeyColumn(-1) #Filtre sur toutes les colonnes
            self.proxy_model.setSourceModel(self.model)
            self.proxy_model.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

            self.tab.setModel(self.proxy_model)
            self.tab.resizeColumnsToContents()

            self.searchBar.textChanged.connect(self.proxy_model.setFilterFixedString)

            self.middleLayout.addWidget(self.searchBar)
            self.middle.setLayout(self.middleLayout)

            ##### Push buttons #####
            fontBouton = QFont()
            fontBouton.setPointSize(12)
            fontBouton.setBold(True)
            fontBouton.setWeight(500)

            self.creer = QPushButton("Créer un dépôt")
            self.creer.clicked.connect(self.creerDepot)
            self.creer.setMaximumSize(QSize(300, 100))
            self.creer.setMinimumSize(QSize(300,40))
            iconCreer = QIcon()
            iconCreer.addPixmap(QPixmap(os.path.join(repertexec,"create.png")), QIcon.Mode.Normal, QIcon.State.Off)
            self.creer.setIcon(iconCreer)

            self.modifier = QPushButton("Modifier un dépôt")
            self.modifier.setFont(fontBouton)
            self.modifier.clicked.connect(self.modifDepot)
            self.modifier.setMaximumSize(QSize(300, 100))
            iconModifier = QIcon()
            iconModifier.addPixmap(QPixmap(os.path.join(repertexec,"edit.png")), QIcon.Mode.Normal, QIcon.State.Off)
            self.modifier.setIcon(iconModifier)

            self.supprimer = QPushButton("Supprimer un dépôt")
            self.supprimer.clicked.connect(self.supprimerDepot)
            self.supprimer.setMaximumSize(QSize(300, 100))
            iconSupprimer = QIcon()
            iconSupprimer.addPixmap(QPixmap(os.path.join(repertexec,"delete.png")), QIcon.Mode.Normal, QIcon.State.Off)
            self.supprimer.setIcon(iconSupprimer)

            self.bandeau = QWidget()
            self.bandeauBoutons = QHBoxLayout()
            self.bandeauBoutons.addWidget(self.creer)
            self.bandeauBoutons.addWidget(self.modifier)
            self.bandeauBoutons.addWidget(self.supprimer)

            self.bandeau.setLayout(self.bandeauBoutons)

            self.boutonValidation = QPushButton("Confirmer modifications")
            self.boutonValidation.setMaximumSize(QSize(300,100))
            self.boutonValidation.setMinimumSize(QSize(300,40))
            self.boutonValidation.clicked.connect(self.saveData)
            iconValider = QIcon()
            iconValider.addPixmap(QPixmap(os.path.join(repertexec,"check.png")), QIcon.Mode.Normal, QIcon.State.Off)
            self.boutonValidation.setIcon(iconValider)

            ##### Comptage nombre de requêtes de changement #####
            self.sheetRequetes = pd.read_excel(req, sheet_name="Requete")
            self.nbrLignes = self.sheetRequetes.shape[0]

            self.traitementRequetes = QPushButton("Requêtes de MAJ: " + str(self.nbrLignes))
            self.traitementRequetes.setMaximumSize(QSize(300,100))
            self.traitementRequetes.clicked.connect(self.traitementRequetesChangement)


            spacer1 = QSpacerItem(100,0)
            spacer2 = QSpacerItem(100,0)

            self.bandeauInf = QWidget()
            self.bandeauInfLayout = QHBoxLayout()

            self.bandeauInfLayout.addItem(spacer1)
            self.bandeauInfLayout.addWidget(self.traitementRequetes)
            self.bandeauInfLayout.addWidget(self.boutonValidation)
            self.bandeauInfLayout.addItem(spacer2)

            self.bandeauInf.setLayout(self.bandeauInfLayout)
            

            ##### Mise en place du widget #####
            self.layoutAdminGUI = QVBoxLayout()
            self.layoutAdminGUI.addWidget(self.topBanner, 0, Qt.AlignmentFlag.AlignHCenter)
            self.layoutAdminGUI.addWidget(self.middle)
            self.layoutAdminGUI.addWidget(self.tab)
            self.layoutAdminGUI.addWidget(self.bandeau)
            self.layoutAdminGUI.addWidget(self.bandeauInf)
            ##### Ajout du bouton réglage pour les SuperAdmins #####
            if self.role == "SuperAdmin":
                
                self.boutonReglage = QPushButton()
                self.boutonReglage.setMinimumSize(60,60)
                self.boutonReglage.setMaximumSize(60,60)
                self.boutonReglage.setIcon(QIcon('reglage.png'))
                self.boutonReglage.setIconSize(self.boutonReglage.size())
                self.boutonReglage.setFixedSize(self.boutonReglage.sizeHint())
                self.boutonReglage.setToolTip('Réglages')
                self.boutonReglage.clicked.connect(self.reglage)

                self.layoutAdminGUI.addWidget(self.boutonReglage, 0, Qt.AlignmentFlag.AlignLeft)

            self.layoutAdminGUI.setSpacing(0)
            self.adminGUI.setLayout(self.layoutAdminGUI)

            ##### Adding the interface to the StackedWidget #####

            self.Stack.addWidget(self.adminGUI)
            self.Stack.setCurrentIndex(1)

        ##### Creation de l'interface non-admin #####
        else:            
            self.adminGUI = QWidget()
            self.sheetRequetes = pd.read_excel(req, sheet_name="Requete")

            ##### Top Banner #####
            self.titre = QLabel("Base de données magasin - Brico Dépôt")
            font = QFont()
            font.setPointSize(28)
            font.setBold(True)
            font.setWeight(75)
            self.titre.setFont(font)
            self.titre.setFrameShape(QFrame.Shape.Box)
            self.titre.setFrameShadow(QFrame.Shadow.Plain)
            self.titre.setLineWidth(5)
            self.titre.setStyleSheet("background-color: rgb(191, 189, 184)")

            self.logo = QLabel("")
            self.logo.setMaximumSize(QSize(210, 130))
            self.logo.setPixmap(QPixmap(os.path.join(repertexec,"logo.png")))
            self.logo.setScaledContents(True)
            self.logo.setStyleSheet("background-color: white")
            self.logo.setStyleSheet("background-color: rgb(191, 189, 184)")
            
            self.infos = QLabel(self.denom + "\n" + self.role)
            font2 = QFont()
            font2.setPointSize(16)
            self.infos.setFont(font2)
            self.infos.setFrameShape(QFrame.Shape.Box)
            self.infos.setFrameShadow(QFrame.Shadow.Plain)
            self.infos.setLineWidth(5)

            self.topBanner = QWidget()

            self.topLayout = QHBoxLayout()
            self.topLayout.setContentsMargins(10, 10, -1, -1)
            self.topLayout.setSpacing(60)
            self.topLayout.addWidget(self.logo)
            self.topLayout.addWidget(self.titre)
            self.topLayout.addWidget(self.infos)
            self.topLayout.addStretch(1)

            self.topBanner.setLayout(self.topLayout)

            ##### Table creation #####
            self.middle = QWidget()
            self.middleLayout = QVBoxLayout()
            self.tab = QTableView()
            self.tab.setModel(self.model)
            self.tab.resizeColumnsToContents()
            self.middleLayout.addWidget(self.tab)
            
            self.middle.setLayout(self.middleLayout)

            ##### Push buttons #####
            self.requete = QPushButton("Demander un changement")
            self.requete.setMaximumSize(300,100)
            self.requete.clicked.connect(self.demandeChangement)

            self.bandeau = QWidget()
            self.bandeauBoutons = QHBoxLayout()
            self.bandeauBoutons.addWidget(self.requete)
            self.bandeau.setLayout(self.bandeauBoutons)

            ##### Mise en place du Widget #####
            self.layoutAdminGUI = QVBoxLayout()
            self.layoutAdminGUI.addWidget(self.topBanner, 0, Qt.AlignmentFlag.AlignHCenter)
            self.layoutAdminGUI.addWidget(self.middle)
            self.layoutAdminGUI.addWidget(self.bandeau)
            self.adminGUI.setLayout(self.layoutAdminGUI)

            ##### Ajout de l'interface dans le stackedWidget #####
            self.Stack.addWidget(self.adminGUI)
            self.Stack.setCurrentIndex(1)

    def center(self):
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()

        qr.moveCenter(cp)
        self.move(qr.topLeft())
        
    def creerDepot(self):
        print("Création")
        self.w = creaWindow()
        self.w.show()

    def modifDepot(self):
        print("Modification")
        if self.role == "SuperAdmin":
            modifWindow.sheet=self.sheet
        else:
            modifWindow.sheet = self.sheet.loc[:, self.lstParam]
        self.w = modifWindow()
        self.w.show()

    def supprimerDepot(self):
        print("Suppression")
        if self.role == "SuperAdmin":
            suppWindow.sheet=self.sheet
        else:

            suppWindow.sheet=self.sheet.loc[:, self.lstParam]
        self.w = suppWindow()
        self.w.show()

    def demandeChangement(self):
        print("Formulaire demande changement")
        demandeChangement.sheet = self.sheet_tri
        responseREQ = File.open_binary(ctx, requete_URL)
        print("Reponse trouvée")
        bytes_file_obj_req = io.BytesIO()
        bytes_file_obj_req.write(responseREQ.content)
        bytes_file_obj_req.seek(0)
        print ("Requêtes chargées")
        req = bytes_file_obj_req
        demandeChangement.sheetOrigines = pd.read_excel(req, sheet_name='Requete')
        self.wDem = demandeChangement()
        self.wDem.gene()
        self.wDem.show()

    def traitementRequetesChangement(self):
        print("Traitement des requetes de changement")
        self.w = traitementDemandeChangement()
        self.w.show()

    def chargerModif(self):
        if self.role == "SuperAdmin":
            self.model = pandasEditableModel(self.sheet)
        elif self.role == "Admin":
            self.model = pandasEditableModel(self.sheet.loc[:, self.lstParam])
        self.proxy_model.setSourceModel(self.model)
        self.tab.setModel(self.proxy_model)
        self.tab.resizeColumnsToContents()
        print("Model chargé")

    def reglage(self):
        print("Réglages")
        self.wReglage = reglages()
        self.wReglage.lstColonnesCheckees = self.lstParam

        self.wReglage.Gui()
        self.wReglage.show()
##TODO checker le save
    def saveData(self):
        print("Saving...")
        self.sheet.to_excel('bddTest.xlsx', index=False)
        
        #Storage du fichier excel temporaire, changer pour le storer uniquement si c'est un utilisateur admin
        self.download_path = os.path.join(tempfile.mkdtemp(), os.path.basename(bdd_URL))
        with open(self.download_path, "wb") as local_file:
            file = ctx.web.get_file_by_server_relative_url(bdd_URL).download(local_file).execute_query()
        print("[Ok] file has been downloaded into: {0}".format(self.download_path))

        self.wb_obj = wb.load_workbook(self.download_path)
        print("Workbook loadé sur openpyxl")

        my_set = set(self.sheet_columns)
        lst_clean = []
        for key in self.d.keys():
            lst_columns = self.d[key].columns.to_list()
            
            for i in range(len(lst_columns)):
                if lst_columns[i] in my_set:
                    lst_clean.append(lst_columns[i])

            self.worksheet = self.wb_obj[key]
            self.worksheet.delete_rows(2, self.worksheet.max_row)

            self.donnees = self.sheet[lst_clean].values.tolist()
            for ligne in self.donnees:
                self.worksheet.append(ligne)

        self.wb_obj.save('BDD2.xlsx')
        
        ########## Upload du fichier sur Sharepoint ##########
        """
        with open(self.download_path, 'rb') as content_file:
            file_content = content_file.read()
        

        file_folder = ctx.web.get_folder_by_server_relative_url("/sites/BricoDepot/Shared%20Documents/Donnees")
        target_file = file_folder.upload_file('BDD2.xlsx', file_content).execute_query()

        print("File hase been uploaded to url: {0}".format(target_file.serverRelativeUrl))
        """
        shutil.rmtree(os.path.dirname(self.download_path))

    def closeEvent(self, event):
        self.w = ''
        if self.w:
            self.w.close()

""" creaWindow : Classe gérant la fenêtre de création de dépôt

    {__init__}: Fonction d'initialisation de la classe

    {automCo}: Création de fenêtre/formulaire en fonction de la bdd

    {confirmerCreation}: Fonction d'ajout du nouveau dépôt dans le dataframe de la fenêtre principale
"""
class creaWindow(QWidget):
    windowslst = {}
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Création d'un dépôt")
        sizeWidth = 870
        sizeHeight = 600
        self.setFixedSize(sizeWidth,sizeHeight)

        font = QFont()
        font.setPointSize(36)
        font.setBold(True)
        font.setWeight(75)

        self.titre = QLabel("Création d'un dépôt")
        self.titre.setFont(font)

        self.VLayout = QVBoxLayout()
        self.VLayout.setSpacing(10)
        self.VLayout.addWidget(self.titre, 0, Qt.AlignmentFlag.AlignHCenter)

        self.sheets = sheet_lst

        if ('Accueil' in self.sheets):
            self.sheets.remove('Accueil')
        if ('BDD' in self.sheets):    
            self.sheets.remove('BDD')

        nbRow = (len(self.sheets) // 4)
        reste = len(self.sheets) % 4

        index = 0

        for iterRow in range(nbRow):
            #Creation Widget de la row
            wid = QWidget()
            nomWid = "Widget" + str(iterRow)
            setattr(self, nomWid, wid)

            #Creation Layout H de la row
            lay = QHBoxLayout()
            nomLay = "layoutH" + str(iterRow)
            setattr(self,nomLay,lay)

            
            for iterCol in range(4):
                pushButton = QPushButton()
                nomPushButton = traitementNom(self.sheets[index])
                contentPushButton = self.sheets[index].replace('_',' ').capitalize()
                setattr(self, nomPushButton, pushButton)
                getattr(self, nomPushButton).setText(contentPushButton)
                getattr(self, nomPushButton).setMaximumSize(200,40)
                getattr(self, nomPushButton).setMinimumSize(200,40)
                self.sheet_name = str(self.sheets[index])
                getattr(self, nomPushButton).clicked.connect(lambda state, x=self.sheet_name, y=False: self.automCo(x, y))

                getattr(self,nomLay).addWidget(getattr(self,nomPushButton))
                self.automCo(self.sheet_name, True)
                index += 1
                

            getattr(self,nomWid).setLayout(getattr(self,nomLay))
            self.VLayout.addWidget(getattr(self,nomWid))
                
        self.layoutReste = QHBoxLayout()
        self.widgetReste = QWidget()
        for i in range(reste):
            pushButton = QPushButton()
            nomPushButton = traitementNom(self.sheets[index])
            contentPushButton = self.sheets[index].replace('_',' ').capitalize()
            self.sheet_name = str(self.sheets[index])
            setattr(self, nomPushButton, pushButton)
            getattr(self, nomPushButton).setText(contentPushButton)
            getattr(self, nomPushButton).setMaximumSize(200,40)
            getattr(self, nomPushButton).setMinimumSize(200,40)
            getattr(self, nomPushButton).clicked.connect(lambda state, x=self.sheet_name, y=False: self.automCo(x, y))
            

            self.layoutReste.addWidget(getattr(self, nomPushButton))
            self.automCo(self.sheet_name, True)
            index += 1

        self.widgetReste.setLayout(self.layoutReste)
        self.VLayout.addWidget(self.widgetReste)

        self.boutonConfirmer = QPushButton("Confirmer la création du dépôt")
        self.boutonConfirmer.clicked.connect(self.confirmerCreation)
        self.boutonConfirmer.setMaximumSize(300,40)
        self.VLayout.addWidget(self.boutonConfirmer, 0, Qt.AlignmentFlag.AlignHCenter)

        self.setLayout(self.VLayout)

    def automCo(self, str, crea):
        if str in self.windowslst:
            self.w = self.windowslst[str]
        else:
            fenetre = varCrea()
            nomFenetre = str
            setattr(self, nomFenetre, fenetre)
            getattr(self,nomFenetre).excel_sheet = str
            getattr(self,nomFenetre).demarrage(str)
            self.windowslst[str] = getattr(self,nomFenetre)
            self.w = getattr(self, nomFenetre)

        if crea == False:
            self.w.show()


    def confirmerCreation(self):
        lstRemove = ['Code BRICO','Code EASIER','Dépôt','Région 2022']

        CreaSheet = pd.DataFrame
        dic = {}

        for i in range (len(self.sheets)):
            self.sheetName = self.sheets[i]
           

            self.lstColonnes = list(pd.read_excel(bdd, sheet_name=self.sheetName).columns.tolist())
            if self.sheetName != 'Liste_depots':
                for j in range (len(lstRemove)):
                    if lstRemove[j] in self.lstColonnes:
                        self.lstColonnes.remove(lstRemove[j])
            
            for j in range (len(self.lstColonnes)):
                nomLineEdit = str(traitementNom(self.lstColonnes[j]))

                self.windows = self.windowslst[self.sheetName]

                valeur = self.windows.getLineEditValue(nomLineEdit)

                

                dic[self.lstColonnes[j]] = valeur

        s = pd.Series(dic)
        newRow = s.to_frame().T

        #Concaténation avec la sheet originale
        CreaSheet = pd.concat([main.sheet, newRow], axis=0)
        CreaSheet = CreaSheet.reset_index(drop=True)

        main.sheet = CreaSheet
        main.chargerModif()
        self.close()
        

    def closeEvent(self, event):
        print("je fais le close event")
        self.w=''
        if self.w:
            self.w.close()

""" varCrea : Classe gérant la création de formulaire de saisie de données pour la créatoin de dépôts

    {__init__}: Fonction d'initialisation de la classe

    {demarrage}: Fonction de création des widgets en fonctions des données de la bdd

    {getLineEditValue}: Fonction pour récupérer la saisie d'un utilisateur
"""
class varCrea(QWidget):
    excel_sheet = ""
    def __init__(self):
        super().__init__()
        self.layout = QGridLayout()
        

    def demarrage(self, sheet_name):
        self.excel_sheet = sheet_name
        self.setWindowTitle(self.excel_sheet)
        self.titre = QLabel("Onglet: " + self.excel_sheet)
        self.layout.addWidget(self.titre, 0, 0, 1, 2, Qt.AlignmentFlag.AlignHCenter)
        self.workSheet = pd.read_excel(bdd, sheet_name=self.excel_sheet)
        self.colonnes = list(self.workSheet.columns.tolist())

        if self.excel_sheet != "Liste_depots":
            lstRemove = ['Code BRICO','Code EASIER','Dépôt','Région 2022']
            for nom in lstRemove:
                if nom in self.colonnes:
                    self.colonnes.remove(nom)

        

        for i in range(len(self.colonnes)):

            label = QLabel()
            nomLabel= "label" + str(traitementNom(self.colonnes[i]))
            setattr(self, nomLabel, label)
            getattr(self, nomLabel).setText(self.colonnes[i])

            lineEdit = QLineEdit()
            nomLineEdit = str(traitementNom(self.colonnes[i]))
            setattr(self, nomLineEdit, lineEdit)
            getattr(self, nomLineEdit).setPlaceholderText(self.colonnes[i])
            getattr(self, nomLineEdit).setMinimumWidth(150)

            self.layout.addWidget(getattr(self,nomLabel), i+1,0)
            self.layout.addWidget(getattr(self,nomLineEdit), i+1,1)

        if len(self.colonnes) > 15:
            self.scroll_area = QScrollArea(self)
            self.scroll_area.setWidgetResizable(True)
            self.scroll_area.setFixedSize(500,380)

            self.widget = QWidget()
            self.widget.setLayout(self.layout)
            self.scroll_area.setWidget(self.widget)
        else:
            self.setLayout(self.layout)

    def getLineEditValue(self, nomLineEdit):
        return getattr(self, nomLineEdit).text()

""" modifWindow : Classe gérant la fenêtre de modification de dépôt

    {__init__}: Fonction d'initialisation de la classe

    {selectionDepot}: Fonction actualisant les données des tableView en fonction de la selection du combobox (première selection)

    {chargementTableau}: Fonction actualisant les données des tableView en fonction de la selection du combobox (seconde selection)

    {choix}: Sauvegarde des modifications dans la dataframe main

    {center}: Fonction pour centrer la fenetre sur l'écran

    {annuler}: Fonction pour quitter la fenêtre
"""
class modifWindow(QWidget):
    sheet = pd.DataFrame

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Modification d'un dépôt")
        self.title = QLabel("Modification d'un dépôt")
        self.setFixedSize(300,80)

        self.stack = QStackedWidget()
                
        self.widget1 = QWidget()
        self.layout1 = QVBoxLayout()

        font = QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)

        self.titre = QLabel("Sélection du dépôt à modifier")
        self.titre.setFont(font)
        self.listedepots = QComboBox()
        self.listedepots.setMinimumHeight(30)
        
        self.listedesdepots = self.sheet["Dépôt"].values.tolist()
        self.listeCodeBRICO = self.sheet["Code BRICO"].values.tolist()

        for i in range (len(self.listedesdepots)):
            self.listedesdepots[i] = str(self.listeCodeBRICO[i]) + "-" + self.listedesdepots[i]

        self.listedepots.addItem('')
        self.listedepots.addItems(self.listedesdepots)
        self.listedepots.currentIndexChanged.connect(self.selectionDepot)



        self.layout1.addWidget(self.titre, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout1.addWidget(self.listedepots, 0, Qt.AlignmentFlag.AlignHCenter)

        self.widget1.setLayout(self.layout1)
        
        self.widget2 = QWidget()
        self.layout2 = QVBoxLayout()

        self.titre2 = QLabel("Sélection du dépôt à modifier")
        self.titre2.setFont(font)
        self.listedepots2 = QComboBox()
        self.listedepots2.setMinimumHeight(30)
        self.listedepots2.addItems(self.listedesdepots)
        self.listedepots2.currentIndexChanged.connect(self.chargementTableau)
        self.affichageDep = QTableView()

        self.boutonConfirmation = QPushButton("Confirmer la modification")
        self.boutonConfirmation.clicked.connect(self.choix)
        self.boutonConfirmation.setMaximumSize(300,100)
        #self.boutonConfirmation.setStyleSheet(styleSheetBouton)

        self.boutonAnnulation = QPushButton("Annuler la modification")
        self.boutonAnnulation.clicked.connect(self.annuler)
        self.boutonAnnulation.setMaximumSize(300,100)
        #self.boutonAnnulation.setStyleSheet(styleSheetBouton)

        self.widgetBouton = QWidget()
        self.layoutBouton = QHBoxLayout()

        self.spacer = QSpacerItem(100,0)
        self.layoutBouton.addWidget(self.boutonAnnulation)
        self.layoutBouton.addItem(self.spacer)
        self.layoutBouton.addWidget(self.boutonConfirmation)
        
        self.widgetBouton.setLayout(self.layoutBouton)
        

        self.layout2.addWidget(self.titre2, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout2.addWidget(self.listedepots2, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout2.addWidget(self.affichageDep)
        self.layout2.addWidget(self.widgetBouton, 0, Qt.AlignmentFlag.AlignHCenter)
        


        self.widget2.setLayout(self.layout2)

        self.mainlayout = QGridLayout()
        self.stack.addWidget(self.widget1)
        self.stack.addWidget(self.widget2)
        self.mainlayout.addWidget(self.stack)
        self.setLayout(self.mainlayout)

    def selectionDepot(self):
        ##Transition vers la seconde fenêtre
        self.listedepots2.setCurrentIndex(self.listedepots.currentIndex()-1)
        ##Charger tableau
        depot = self.listedepots2.currentText()
        depot = str(depot.split('-')[1])

        self.sheet_tri = self.sheet.loc[self.sheet['Dépôt'] == depot]
        self.sheet_tri_index = self.sheet.loc[self.sheet['Dépôt'] == depot].index[0]

        self.modelModif = pandasEditableModel(self.sheet_tri)
        self.affichageDep.setModel(self.modelModif)
        self.affichageDep.resizeColumnsToContents()

        self.stack.setCurrentIndex(1)    
        self.setFixedSize(1020,540)     
        self.center()      

    def chargementTableau(self):
        depot = self.listedepots2.currentText()
        depot = str(depot.split('-')[1])

        self.sheet_tri = self.sheet.loc[self.sheet['Dépôt'] == depot]
        self.sheet_tri_index = self.sheet.loc[self.sheet['Dépôt'] == depot].index[0]

        self.modelModif = pandasEditableModel(self.sheet_tri)
        self.affichageDep.setModel(self.modelModif)
        self.affichageDep.resizeColumnsToContents()

    def choix(self):
    ##Application des modifications sur la sheet de la classe
        rowIndex = self.sheet_tri_index
        listHeaders = self.sheet.columns.values.tolist()

        for i in range(len(listHeaders)):
            main.sheet[listHeaders[i]][rowIndex] = self.sheet_tri[listHeaders[i]][rowIndex]
                
    ##Transmission de la nouvelle sheet à la fenêtre principale
        main.chargerModif()

        self.close()

    def center(self):
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()

        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def annuler(self):
        self.close()

""" suppWindow : Classe gérant la fenêtre de suppression de dépôt

    {__init__}: Fonction d'initialisation de la classe

    {selectionDepot}: Fonction actualisant les données des tableView en fonction de la selection du combobox (première selection)

    {chargementTableau}: Fonction actualisant les données des tableView en fonction de la selection du combobox (seconde selection) 

    {suppression}: Suppression du dépôt dans la dataframe main

    {center}: Fonction pour centrer la fenêtre sur l'écran

    {annuler}: Fonction pour quitter la fenêtre
"""
class suppWindow(QWidget):
    sheet = pd.DataFrame
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Suppression d'un dépôt")

        font = QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        self.title = QLabel("Suppression d'un dépôt")
        

        self.setFixedSize(300,80)

        self.stack = QStackedWidget()
        self.i= 5

        self.widget1 = QWidget()
        self.layout1 = QVBoxLayout()

        self.titre = QLabel("Sélection du dépôt à supprimer")
        self.titre.setFont(font)
        self.listedepots = QComboBox()
        self.listedepots.setMinimumHeight(30)
        self.listedesdepots = self.sheet["Dépôt"].values.tolist()
        self.listeCodeBRICO = self.sheet["Code BRICO"].values.tolist()

        for i in range (len(self.listedesdepots)):
            self.listedesdepots[i] = str(self.listeCodeBRICO[i]) + "-" + str(self.listedesdepots[i])

        self.listedepots.addItem('')
        self.listedepots.addItems(self.listedesdepots)
        self.listedepots.currentIndexChanged.connect(self.selectionDepot)

        self.layout1.addWidget(self.titre, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout1.addWidget(self.listedepots, 0, Qt.AlignmentFlag.AlignHCenter)

        self.widget1.setLayout(self.layout1)
        

        self.widget2 = QWidget()
        self.layout2 = QVBoxLayout()
        self.titre2 = QLabel("Sélection du dépôt à supprimer")
        self.titre2.setFont(font)
        self.listedepots2 = QComboBox()
        self.listedepots2.setMinimumHeight(30)
        self.listedepots2.addItems(self.listedesdepots)
        self.listedepots2.currentIndexChanged.connect(self.chargementTableau)
        self.affichageDep = QTableView()

        self.boutonConfirmation = QPushButton("Supprimer le dépôt")
        self.boutonConfirmation.clicked.connect(self.suppression)
        self.boutonConfirmation.setMaximumSize(300,100)
        #self.boutonConfirmation.setStyleSheet(styleSheetBouton)

        self.boutonAnnulation = QPushButton("Annuler")
        self.boutonAnnulation.clicked.connect(self.annuler)
        self.boutonAnnulation.setMaximumSize(300,100)
        #self.boutonAnnulation.setStyleSheet(styleSheetBouton)

        self.widgetBouton = QWidget()
        self.layoutBouton = QHBoxLayout()

        self.space = QSpacerItem(100,0)
        self.layoutBouton.addWidget(self.boutonAnnulation)
        self.layoutBouton.addItem(self.space)
        self.layoutBouton.addWidget(self.boutonConfirmation)

        self.widgetBouton.setLayout(self.layoutBouton)

        #, Qt.AlignmentFlag.AlignCenter
        self.layout2.addWidget(self.titre2, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout2.addWidget(self.listedepots2, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout2.addWidget(self.affichageDep)
        self.layout2.addWidget(self.widgetBouton)
        

        self.widget2.setLayout(self.layout2)

        self.stack.addWidget(self.widget1)
        self.stack.addWidget(self.widget2)

        self.mainlayout = QGridLayout()
        self.mainlayout.addWidget(self.stack)
        self.setLayout(self.mainlayout)
    
    def selectionDepot(self):
        ##Transition vers la seconde fenêtre
        self.listedepots2.setCurrentIndex(self.listedepots.currentIndex()-1)

        depot = self.listedepots2.currentText()
        depot = str(depot.split('-')[1])

        self.sheet_tri = self.sheet.loc[self.sheet['Dépôt'] == depot]
        self.sheet_tri_index = self.sheet.loc[self.sheet['Dépôt'] == depot].index[0]

        self.modelModif = pandasModel(self.sheet_tri)
        self.affichageDep.setModel(self.modelModif)
        self.affichageDep.resizeColumnsToContents()

        self.stack.setCurrentIndex(1)    
        self.setFixedSize(720,440)
        self.center()
        
    def chargementTableau(self):
        print("Chargement")

        depot = self.listedepots2.currentText()
        depot = str(depot.split('-')[1])

        self.sheet_tri = self.sheet.loc[self.sheet['Dépôt'] == depot]
        self.sheet_tri_index = self.sheet.loc[self.sheet['Dépôt'] == depot].index[0]

        self.modelModif = pandasModel(self.sheet_tri)
        self.affichageDep.setModel(self.modelModif)
        self.affichageDep.resizeColumnsToContents()
        self.header = self.affichageDep.horizontalHeader()
        self.header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        print("Chargement fini")

    def suppression(self):
        rowIndex = self.sheet_tri_index
        
        main.sheet = main.sheet.drop(rowIndex)

        main.chargerModif()

        self.close()

        ##TODO: Faire une fenetre de confirmation

    def center(self):
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()

        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def annuler(self):
        self.close()

""" demandeChangement : Classe gérant la fenêtre de demande de mise à jour

    {__init__}: Fonction d'initialisation de la classe

    {transmettre}: Fonction pour valider la demande de mise à jour. Etudie les potentiels conflits et les repertories pour traitement des conflits

    {gene}: Fonction pour créer la liste qui alimente la combobox

    {select}: Fonction pour charger les données du dépôt sélectionné

    {annulerF}: Fonction pour fermer la fenêtre
"""
class demandeChangement(QWidget):
    sheet = pd.DataFrame
    sheetOrigines = pd.DataFrame

    def __init__(self):
        super().__init__()
        self.setFixedSize(1280,720)
        
        self.dataframereq = pd.read_excel(req, sheet_name='Requete')
        self.setWindowTitle("Demande de changement de données")

        self.layout = QVBoxLayout()

        self.Titre = QLabel("Demande de changement de données")
        font = QFont()
        font.setPointSize(28)
        font.setBold(True)
        font.setWeight(75)
        self.Titre.setFont(font)

        self.listeCodes = QComboBox()
        self.listeCodes.setMinimumHeight(30)
        self.listeCodes.currentIndexChanged.connect(self.select)

        self.texte = QLabel("Modifier les données dans le tableau ci-dessous")
        font.setPointSize(14)
        self.texte.setFont(font)

        self.table = QTableView()
        self.model = pandasEditableModel(self.sheet)
        self.table.setModel(self.model)
        self.table.resizeColumnsToContents()

        self.bandeauBouton = QWidget()
        self.bandeauBoutonsLayout = QHBoxLayout()

        self.valider = QPushButton("Valider la demande de\nchangement de données")
        self.valider.setMaximumSize(400,100)
        self.valider.clicked.connect(self.transmettre)

        self.annuler = QPushButton("Annuler la demande de\nchangement de données")
        self.annuler.setMaximumSize(400,100)
        self.annuler.clicked.connect(self.annulerF)

        self.space = QSpacerItem(100,0)
        self.space1 = QSpacerItem(100,0)

        self.bandeauBoutonsLayout.addItem(self.space)
        self.bandeauBoutonsLayout.addWidget(self.annuler)
        self.bandeauBoutonsLayout.addWidget(self.valider)
        self.bandeauBoutonsLayout.addItem(self.space1)

        self.bandeauBouton.setLayout(self.bandeauBoutonsLayout)
        
        self.layout.addWidget(self.Titre, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout.addWidget(self.listeCodes, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout.addWidget(self.texte, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout.addWidget(self.table)
        self.layout.addWidget(self.bandeauBouton)
        

        self.layout.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)

        self.setLayout(self.layout)

    def transmettre(self):
        self.sheet = self.sheet.reset_index(drop=True)

        responseREQ = File.open_binary(ctx, requete_URL)
        print("Reponse trouvée")
        bytes_file_obj_req = io.BytesIO()
        bytes_file_obj_req.write(responseREQ.content)
        bytes_file_obj_req.seek(0)
        print ("Requêtes chargées")
        req = bytes_file_obj_req
        self.dataframereq = pd.read_excel(req, sheet_name='Requete')

        self.rowDemande['Utilisateur'] = main.denom
        self.rowDemande['Date demande'] = datetime.datetime.now().strftime('%Y-%m-%d')

        self.lstDonneesSup = ['Utilisateur', 'Date demande']

        lstCodeBrico = self.dataframereq['Code BRICO'].tolist()

        #Création de la sheet de conflit
        self.sheetC = pd.DataFrame(columns=self.sheet.columns)
        lstCodesConflits = []
        lstCleanConflits = []

        conflit = False
        if self.rowDemande.shape[0] == 1:
            
            codeATrouver = self.rowDemande['Code BRICO'][0]
            if codeATrouver in lstCodeBrico:
                print("Conflit pour le code " + str(codeATrouver))
                index = self.rowDemande.loc[self.rowDemande['Code BRICO'] == codeATrouver].index[0]
                valeur = self.rowDemande.loc[index, 'Dépôt']
                lstCodesConflits.append(str(codeATrouver) + "-" + str(valeur))
                lstCleanConflits.append(codeATrouver)
                conflit = True
        elif self.rowDemande.shape[0] > 1:
            for i in range(self.sheet.shape[0]):
                codeATrouver = self.sheet['Code BRICO'][i]
                if codeATrouver in lstCodeBrico:
                    print("Conflit pour le code " + str(codeATrouver))
                    conflit = True
                    index = self.sheet.loc[self.sheet['Code BRICO'] == codeATrouver].index[0]
                    valeur = self.sheet.loc[index, 'Dépôt']
                    lstCodesConflits.append(str(codeATrouver) + "-" + str(valeur))
                    lstCleanConflits.append(codeATrouver)

        if conflit == True:
            print("Conflit MAJ, ouverture gestionnaire")
            self.gestionConflit = confirmDemande()

            self.gestionConflit.sheetDemande = self.rowDemande
            self.gestionConflit.sheetConflit = self.dataframereq[self.dataframereq['Code BRICO'].isin(lstCleanConflits)]
        
            self.gestionConflit.lstCodes = lstCodesConflits
            self.gestionConflit.lstClean = lstCleanConflits

            self.gestionConflit.geneSheet()
            self.gestionConflit.show()
            
        else:
            self.listeParamCheck = listeParam
            self.listeParamCheck.append('Utilisateur')
            self.listeParamCheck.append('Date demande')
            self.dfCompletee = pd.DataFrame(columns=self.dataframereq.columns)
            for i in range(self.rowDemande.shape[0]):
                self.dfCompletee.loc[len(self.dfCompletee)] = [None] * len(self.dfCompletee.columns)
                codeBrico = self.rowDemande['Code BRICO'][i]
                indexdfReq = main.sheet.loc[main.sheet['Code BRICO'] == codeBrico].index[0]
                print(indexdfReq)
                for j in range(self.dfCompletee.shape[1]):
                    nomColonne = self.dfCompletee.columns[j]
                    if nomColonne in self.listeParamCheck:
                        self.dfCompletee[nomColonne][i] = self.rowDemande[nomColonne][i]
                    else:
                        self.dfCompletee[nomColonne][i] = main.sheet[nomColonne][indexdfReq]
            
            
        
            #Upload sur sharepoint
            download_path_req = os.path.join(tempfile.mkdtemp(), os.path.basename(bdd_URL))
            with open(download_path_req, "wb") as local_file:
                file = ctx.web.get_file_by_server_relative_url(requete_URL).download(local_file).execute_query()
            print("[Ok] file has been downloaded into: {0}".format(download_path_req))

            self.newDF = pd.DataFrame
            self.dataframereq = pd.read_excel(download_path_req, sheet_name='Requete')
            self.newDF = pd.concat([self.dataframereq,self.dfCompletee], axis=0)
            self.newDF = self.newDF.reset_index(drop=True)        

            with pd.ExcelWriter(download_path_req, mode='a', if_sheet_exists='replace') as writer:
                self.newDF.to_excel(writer, sheet_name='Requete', index=False)


            with open(download_path_req, 'rb') as content_file:
                file_content = content_file.read()
            

            file_folder = ctx.web.get_folder_by_server_relative_url("/sites/BricoDepot/Shared%20Documents/Donnees")
            target_file = file_folder.upload_file('REQ.xlsx', file_content).execute_query()

            print("File hase been uploaded to url: {0}".format(target_file.serverRelativeUrl))

            #Suppression du dossier temp
            shutil.rmtree(os.path.dirname(download_path_req))
            
            #Affichage message box, changement confirmé et transmis
            QMessageBox.information(self, 'Succès', 'Requête transmise')
            
            
            self.close()

    def gene(self):
        self.lstCodesBrico = self.sheet['Code BRICO'].tolist()
        self.lstDepots = self.sheet['Dépôt'].tolist()
        for i in range(len(self.lstCodesBrico)):
            self.lstCodesBrico[i] = str(self.lstCodesBrico[i]) + "-" + self.lstDepots[i]
        self.listeCodes.addItems(self.lstCodesBrico)

    def select(self):
        depot = self.listeCodes.currentText()
        depot = str(depot.split('-')[0])
        
        self.rowDemande = self.sheet.loc[self.sheet['Code BRICO'] == int(depot)]
        self.rowDemande = self.rowDemande.reset_index(drop=True)

        self.model = pandasEditableModel(self.rowDemande)
        
        self.table.setModel(self.model)
        self.table.resizeColumnsToContents()

    def annulerF(self):
        self.close()

""" confirmDemande : Classe gérant la fenêtre d'étude de conflits

    {__init__}: Fonction d'initialisation de la classe

    {geneSheet}: Fonction générant les 2 dataframes pour comparer les conflits dans des tableViews

    {annuler}: Fonction pour quitter la fenêtre

    {confremplacement}: MessageBox pour confirmer le remplacement de la requête actuelle par la requête utilisateur

    {remplacement}: Fonction remplacant la requête actuelle par la requête utilisateur après confirmation

    {selectConflit}: Fonction pour actualiser les tableView en fonction de la sélection de la combobox
"""
class confirmDemande(QWidget):
    sheetDemande = pd.DataFrame
    sheetConflit = pd.DataFrame

    lstCodes = []
    lstClean = []

    def __init__(self):
        super().__init__()
        self.setFixedSize(1180,670)
        self.setWindowTitle("Conflit MAJ")

        self.layout = QVBoxLayout()

        font = QFont()
        font.setPointSize(28)
        font.setBold(True)
        font.setWeight(75)

        self.Titre = QLabel("Conflit lors de la demande de MAJ")
        self.Titre.setFont(font)

        self.details = QLabel("Votre demande est en conflit avec une ou plusieurs autres demandes de mise à jour")
        font.setPointSize(14)
        self.details.setFont(font)


        self.listeCodes = QComboBox()
        self.listeCodes.setMinimumHeight(30)
        self.listeCodes.currentIndexChanged.connect(self.selectConflit)


        self.tableNew = QTableView()
        self.tableOrigin = QTableView()

        ###Push buttons###
        self.stretch = QSpacerItem(50,0)
        self.stretch1 = QSpacerItem(50,0)

        self.boutonAnnuler = QPushButton("Annuler")
        self.boutonAnnuler.setMaximumSize(320,100)
        self.boutonAnnuler.setMinimumSize(300,40)
        self.boutonAnnuler.clicked.connect(self.annuler)

        self.boutonRemplacement = QPushButton("Remplacer avec la requête actuelle")
        self.boutonRemplacement.setMaximumSize(320,100)
        self.boutonRemplacement.setMinimumSize(300,40)
        self.boutonRemplacement.clicked.connect(self.confremplacement)

        self.bandeauBouton = QWidget()
        self.layoutBandeauBouton = QHBoxLayout()

        self.layoutBandeauBouton.addWidget(self.boutonAnnuler)
        self.layoutBandeauBouton.addItem(self.stretch)
        self.layoutBandeauBouton.addWidget(self.boutonRemplacement)


        self.bandeauBouton.setLayout(self.layoutBandeauBouton)


        self.layout.addWidget(self.Titre, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout.addWidget(self.details, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout.addWidget(self.listeCodes, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout.addWidget(self.tableNew)
        self.layout.addWidget(self.tableOrigin)
        self.layout.addWidget(self.bandeauBouton, 0, Qt.AlignmentFlag.AlignHCenter)

        self.setLayout(self.layout)

    def geneSheet(self):

        self.sheetConflit = self.sheetConflit[self.sheetConflit['Code BRICO'].isin(self.lstClean)]
        self.sheetDemande = self.sheetDemande[self.sheetDemande['Code BRICO'].isin(self.lstClean)]
        self.model = pandasModel(self.sheetDemande)
        self.model2 = pandasModel(self.sheetConflit)

        self.tableNew.setModel(self.model)
        self.tableNew.resizeColumnsToContents()

        self.tableOrigin.setModel(self.model2)
        self.tableOrigin.resizeColumnsToContents()

        self.listeCodes.addItems(self.lstCodes)

    def annuler(self):
        self.close()

    def confremplacement(self):
        confirm_box = QMessageBox()
        confirm_box.setText("Remplacer la requête existante par votre requête ?")
        
        fontMessage = QFont()
        fontMessage.setPointSize(14)
        confirm_box.setFont(fontMessage)
        confirm_box.setWindowTitle("Confirmation")

        confirm_button = QPushButton("Confirmer")
        cancel_button = QPushButton("Annuler")
        
        confirm_box.addButton(confirm_button, QMessageBox.ButtonRole.AcceptRole)
        confirm_box.addButton(cancel_button, QMessageBox.ButtonRole.RejectRole)
        
        confirm_box.exec()

        if confirm_box.clickedButton() == confirm_button:
            print("L'utilisateur a confirmé.")
            self.remplacement()
        else:
            print("L'utilisateur a annulé.")
  
    def remplacement(self):
        responseREQ = File.open_binary(ctx, requete_URL)
        print("Reponse trouvée")
        bytes_file_obj_req = io.BytesIO()
        bytes_file_obj_req.write(responseREQ.content)
        bytes_file_obj_req.seek(0)
        print ("Requêtes chargées")
        req = bytes_file_obj_req
        main.sheetRequetes = pd.read_excel(req, sheet_name='Requete')

        depot = str(self.listeCodes.currentText().split('-')[0])
        print(depot)
        indexReq = main.sheetRequetes.loc[main.sheetRequetes['Code BRICO'] == int(depot)].index[0]

        #Remplacement des données requêtes
        self.listeParamCheck = listeParam
        self.listeParamCheck.append('Utilisateur')
        self.listeParamCheck.append('Date demande')

        self.dfCompletee = pd.DataFrame(columns=main.sheetRequetes.columns)
        for i in range(self.sheetDemande.shape[0]):
            self.dfCompletee.loc[len(self.dfCompletee)] = [None] * len(self.dfCompletee.columns)
            codeBrico = self.sheetDemande['Code BRICO'][i]
            indexdfReq = main.sheet.loc[main.sheet['Code BRICO'] == codeBrico].index[0]
            for j in range(self.dfCompletee.shape[1]):
                nomColonne = self.dfCompletee.columns[j]
                if nomColonne in self.listeParamCheck:
                    self.dfCompletee[nomColonne][i] = self.sheetDemande[nomColonne][i]
                else:
                    self.dfCompletee[nomColonne][i] = main.sheet[nomColonne][indexdfReq]

        main.sheetRequetes.loc[indexReq] = self.dfCompletee.loc[0]
        
        #Upload sur sharepoint

        download_path_req = os.path.join(tempfile.mkdtemp(), os.path.basename(bdd_URL))
        with open(download_path_req, "wb") as local_file:
            file = ctx.web.get_file_by_server_relative_url(requete_URL).download(local_file).execute_query()
        print("[Ok] file has been downloaded into: {0}".format(download_path_req))

        with pd.ExcelWriter(download_path_req, mode='a', if_sheet_exists='replace') as writer:
            main.sheetRequetes.to_excel(writer, sheet_name='Requete', index=False)


        with open(download_path_req, 'rb') as content_file:
            file_content = content_file.read()
            

        file_folder = ctx.web.get_folder_by_server_relative_url("/sites/BricoDepot/Shared%20Documents/Donnees")
        target_file = file_folder.upload_file('REQ.xlsx', file_content).execute_query()

        print("File hase been uploaded to url: {0}".format(target_file.serverRelativeUrl))

        #Suppression du dossier temp
        shutil.rmtree(os.path.dirname(download_path_req))
            
        #Affichage message box, changement confirmé et transmis
        QMessageBox.information(self, 'Succès', 'Requête remplacée')
        self.close()
                
    def selectConflit(self):
        depot = self.listeCodes.currentText()
        depot = str(depot.split('-')[0])

        self.rowDemande = self.sheetDemande.loc[self.sheetDemande['Code BRICO'] == int(depot)]
        self.rowDemande = self.rowDemande.reset_index(drop=True)
        self.rowConflit = self.sheetConflit.loc[self.sheetConflit['Code BRICO'] == int(depot)]
        self.rowConflit = self.rowConflit.reset_index(drop=True)

        self.model = pandasModel(self.rowDemande)
        self.model2 = pandasModel(self.rowConflit)

        self.tableNew.setModel(self.model)
        self.tableNew.resizeColumnsToContents()

        self.tableOrigin.setModel(self.model2)
        self.tableOrigin.resizeColumnsToContents()

""" traitementDemandeChangement : Classe gérant la fenêtre de traitement des demandes de MAJ

    {__init__}: Fonction d'initialisation de la classe

    {selectionDepot}: Première selection dans la combobox
    
    {selectionDepot2}: Selection suivante dans la combobox. Charge les données actuelles et la requete et analyse les données qui diffèrent entre les deux

    {valider}: Valide la requête et modifie la dataframe main puis sauvegarde les modifications sur sharepoint
 
    {retirer}: Retire la requête

    {center}: Fonction pour centrer la fenetre sur l'écran

    {annuler}: Fonction pour quitter la fenêtre
"""
class traitementDemandeChangement(QWidget):

    def __init__(self):
        super().__init__()

        ##### Gathering the requests #####
        self.sheet = main.sheet
        self.setWindowTitle("Traitement demandes de MAJ")

        #self.dfRequete = pd.read_excel(req, sheet_name='Requete')
        responseREQ = File.open_binary(ctx, requete_URL)
        print("Reponse trouvée")
        bytes_file_obj_req = io.BytesIO()
        bytes_file_obj_req.write(responseREQ.content)
        bytes_file_obj_req.seek(0)
        print ("Requêtes chargées")
        req = bytes_file_obj_req
        self.dfRequete = pd.read_excel(req, sheet_name='Requete')
        #self.dfRequete = main.sheetRequetes 
        self.dfRequete.insert(0, 'TypeLigne', "")
        self.dfRequete['TypeLigne'] = ['Requête de changement'] * len(self.dfRequete.index)


        self.listedesdepotsReq = self.dfRequete["Dépôt"].values.tolist()
        self.listeCodeBRICOReq = self.dfRequete["Code BRICO"].values.tolist()

        for i in range (len(self.listedesdepotsReq)):
            self.listedesdepotsReq[i] = str(self.listeCodeBRICOReq[i]) + "-" + str(self.listedesdepotsReq[i])


        self.setFixedSize(300,80)

        self.stack = QStackedWidget()
        
        ##### WIDGET 1 #####
        self.widget1 = QWidget()
        self.layout1 = QVBoxLayout()

        font = QFont()
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)

        self.titre1 = QLabel("Liste des requêtes à traiter")
        self.titre1.setFont(font)
        self.comboBox1 = QComboBox()
        self.comboBox1.setMinimumHeight(30)
        self.comboBox1.addItem('')
        self.comboBox1.addItems(self.listedesdepotsReq)
        self.comboBox1.currentIndexChanged.connect(self.selectionDepot)

        self.layout1.addWidget(self.titre1)
        self.layout1.addWidget(self.comboBox1)

        self.widget1.setLayout(self.layout1)

        ##### WIDGET 2 #####
        self.widget2 = QWidget()
        self.layout2 = QVBoxLayout()

        self.titre2 = QLabel("Dépôt selectionné")
        self.titre2.setFont(font)
        self.tableModif = QTableView()
        self.comboBox2 = QComboBox()
        self.comboBox2.setMinimumHeight(30)
        self.comboBox2.addItems(self.listedesdepotsReq)
        self.comboBox2.currentIndexChanged.connect(self.selectionDepot2)

        self.boutonValidation = QPushButton("Valider la requête")
        self.boutonValidation.clicked.connect(self.valider)
        self.boutonValidation.setMaximumSize(300,100)
        self.boutonValidation.setMinimumWidth(200)

        self.boutonRetirer = QPushButton("Retirer la requête")
        self.boutonRetirer.clicked.connect(self.retirer)
        self.boutonRetirer.setMaximumSize(300,100)
        self.boutonRetirer.setMinimumWidth(200)

        self.boutonAnnuler = QPushButton("Annuler")
        self.boutonAnnuler.clicked.connect(self.annuler)
        self.boutonAnnuler.setMaximumSize(300,100)
        self.boutonAnnuler.setMinimumWidth(200)

        self.bandeauBouton = QWidget()
        self.bandeauBoutonLayout = QHBoxLayout()

        self.bandeauBoutonLayout.addWidget(self.boutonAnnuler)
        self.bandeauBoutonLayout.addWidget(self.boutonRetirer)
        self.bandeauBoutonLayout.addWidget(self.boutonValidation)
        self.bandeauBouton.setLayout(self.bandeauBoutonLayout)

        self.layout2.addWidget(self.titre2, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout2.addWidget(self.comboBox2, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout2.addWidget(self.tableModif)
        self.layout2.addWidget(self.bandeauBouton, 0, Qt.AlignmentFlag.AlignHCenter)

        self.widget2.setLayout(self.layout2)

        ##### Stack #####

        self.stack.addWidget(self.widget1)
        self.stack.addWidget(self.widget2)

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.stack)
        self.setLayout(self.layout)
    
    def selectionDepot(self):

        self.setFixedSize(720,440)
        self.center()
        if self.stack.currentIndex != 1:
            self.stack.setCurrentIndex(1)
            self.comboBox2.setCurrentIndex(self.comboBox1.currentIndex() - 1)
        self.selectionDepot2()

    def selectionDepot2(self):
        #Création de la sheet
        depot = self.comboBox2.currentText()
        depot = str(depot.split('-')[1])

        #Ligne actuelle
        self.actualdata = self.sheet.loc[self.sheet['Dépôt'] == depot]
        self.actualdata.insert(0, 'TypeLigne', ['Données actuelles'])

        #Ligne de changment
        self.sheetRequete = self.dfRequete.loc[self.dfRequete['Dépôt'] == depot]

        self.sheet_tri = pd.concat([self.actualdata, self.sheetRequete], axis=0)
        self.sheet_tri = self.sheet_tri.reset_index(drop=True)

        self.lstParam = listeParam
        self.lstParam.append('Utilisateur')
        self.lstParam.append('Date demande')
        self.sheet_tri = self.sheet_tri.loc[:, listeParam]

        #Chargement table
        self.model = pandasModel(self.sheet_tri)
        
        self.tableModif.setModel(self.model)
    
        self.tableModif.resizeColumnsToContents()
        self.header = self.tableModif.horizontalHeader()
        self.header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)

        #Surligner en rouge 
        colonnes_dif = []
        modifs = []
        for colonne in range(len(self.sheet_tri.columns.tolist())):
            if str(self.sheet_tri.iloc[0, colonne]) != str(self.sheet_tri.iloc[1, colonne]):
                colonnes_dif.append(colonne)
                modification = str(self.sheet_tri.iloc[0, colonne]) + " devient : " + str(self.sheet_tri.iloc[1, colonne])
                modifs.append(modification)
        
        print(colonnes_dif)
        print(modifs)

        for i in range(len(colonnes_dif)):
            self.model.change_color(1, colonnes_dif[i], QBrush(Qt.GlobalColor.red))

    def valider(self):
        print("Valider la requête")

        lst_colonnes = main.sheet.columns.tolist()
        self.codeBrico = self.sheet_tri.iloc[1]["Code BRICO"]
        self.depot = self.sheet_tri.iloc[1]["Dépôt"]
        self.utilisateur = self.sheet_tri.iloc[1]["Utilisateur"]
        self.dateDemande = self.sheet_tri.iloc[1]["Date demande"]


        #Index de la ligne à changer
        index_ligne = self.dfRequete.loc[(self.dfRequete['Code BRICO'] == self.codeBrico)].index[0]
        indexMain = main.sheet.loc[(main.sheet['Code BRICO'] == self.codeBrico)].index[0]

        for i in range(len(lst_colonnes)):
            main.sheet[lst_colonnes[i]][indexMain] = self.dfRequete[lst_colonnes[i]][index_ligne]

        main.chargerModif()

        #La requête est passée, supprimer la requete
        self.dfRequete = self.dfRequete.drop(index_ligne)
        self.dfRequete.reset_index(drop=True)

        self.dfReqWrite = self.dfRequete

        download_path_req = os.path.join(tempfile.mkdtemp(), os.path.basename(bdd_URL))
        with open(download_path_req, "wb") as local_file:
            file = ctx.web.get_file_by_server_relative_url(requete_URL).download(local_file).execute_query()
        print("[Ok] file has been downloaded into: {0}".format(download_path_req))

        with pd.ExcelWriter(download_path_req, mode='a', if_sheet_exists='replace') as writer:
            self.dfReqWrite.to_excel(writer, sheet_name='Requete', index=False)


        with open(download_path_req, 'rb') as content_file:
            file_content = content_file.read()
        

        file_folder = ctx.web.get_folder_by_server_relative_url("/sites/BricoDepot/Shared%20Documents/Donnees")
        target_file = file_folder.upload_file('REQ.xlsx', file_content).execute_query()

        print("File hase been uploaded to url: {0}".format(target_file.serverRelativeUrl))

        #Suppression du dossier temp
        shutil.rmtree(os.path.dirname(download_path_req))
        
        #Affichage message box, changement confirmé et transmis
        QMessageBox.information(self, 'Succès', 'Requête acceptée')

        if self.dfRequete.empty:
            QMessageBox.information(self, 'Succès', 'Plus de requêtes à traiter, fermeture...')
            self.close()
            return
        
        curInd = self.comboBox2.currentIndex()
        self.comboBox2.removeItem(curInd)
        self.comboBox2.setCurrentIndex(1)

    ##TODO: Refaire le test avec 1 requete et plusieurs requetes
    ##TODO: Mettre a jour dans la main window, retirer la requete aussi
    
    def retirer(self):

        self.codeBrico = self.sheet_tri.iloc[1]["Code BRICO"]
        self.depot = self.sheet_tri.iloc[1]["Dépôt"]
        self.utilisateur = self.sheet_tri.iloc[1]["Utilisateur"]
        self.dateDemande = self.sheet_tri.iloc[1]["Date demande"]

        index_ligne = self.dfRequete.loc[(self.dfRequete['Code BRICO'] == self.codeBrico)].index[0]

        self.dfRequete = self.dfRequete.drop(index_ligne)
        self.dfRequete.reset_index(drop=True)

        #Maj Sharepoint
        self.dfReqWrite = self.dfRequete.drop("TypeLigne", axis = 1)
        download_path_req = os.path.join(tempfile.mkdtemp(), os.path.basename(bdd_URL))
        with open(download_path_req, "wb") as local_file:
            file = ctx.web.get_file_by_server_relative_url(requete_URL).download(local_file).execute_query()
        print("[Ok] file has been downloaded into: {0}".format(download_path_req))

        with pd.ExcelWriter(download_path_req, mode='a', if_sheet_exists='replace') as writer:
            self.dfReqWrite.to_excel(writer, sheet_name='Requete', index=False)


        with open(download_path_req, 'rb') as content_file:
            file_content = content_file.read()
        

        file_folder = ctx.web.get_folder_by_server_relative_url("/sites/BricoDepot/Shared%20Documents/Donnees")
        target_file = file_folder.upload_file('REQ.xlsx', file_content).execute_query()

        print("File hase been uploaded to url: {0}".format(target_file.serverRelativeUrl))

        #Suppression du dossier temp
        shutil.rmtree(os.path.dirname(download_path_req))
        
        #Affichage message box, changement confirmé et transmis
        QMessageBox.information(self, 'Succès', 'Requête supprimée')

        responseREQ = File.open_binary(ctx, requete_URL)
        print("Reponse trouvée")
        bytes_file_obj_req = io.BytesIO()
        bytes_file_obj_req.write(responseREQ.content)
        bytes_file_obj_req.seek(0)
        print ("Requêtes chargées")
        req = bytes_file_obj_req
        self.dfRequete = pd.read_excel(req, sheet_name='Requete')
        self.dfRequete.insert(0, 'TypeLigne', "")
        self.dfRequete['TypeLigne'] = ['Requête de changement'] * len(self.dfRequete.index)

        if self.dfRequete.empty:
            QMessageBox.information(self, 'Succès', 'Plus de requêtes à traiter, fermeture...')
            main.traitementRequetes.setText("Requêtes de MAJ: " + str(0))
            self.close()
            return
        
        curInd = self.comboBox2.currentIndex()
        self.comboBox2.removeItem(curInd)
        print(curInd)
        if curInd == 0:
            self.comboBox2.setCurrentIndex(0)
        else:
            self.comboBox2.setCurrentIndex(curInd-1)

        self.nbrLignes = self.dfRequete.shape[0]

        main.traitementRequetes.setText("Requêtes de MAJ: " + str(self.nbrLignes))
        main.sheetRequetes = self.dfRequete

    def center(self):
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()

        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def annuler(self):
        self.close()   

""" reglages : Classe gérant la fenêtre des réglages

    {__init__}: Fonction d'initialisation de la classe

    {gui}: Fonction pour initialiser l'interface de la fenêtre réglage

    {save}: Fonction de sauvegarde des réglages
"""
class reglages(QWidget):
    lstColonnesComplete = []
    lstColonnesCheckees = []

    lstcheckcrees = []
    
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Réglage")
        #self.setFixedSize(400,550)

    def Gui(self):
        self.sheetLst = sheet_lst

        self.layout = QVBoxLayout()

        self.font = QFont()
        self.font.setPointSize(14)
        self.font.setBold(True)

        self.font2 = QFont()
        self.font2.setPointSize(12)
        self.font2.setBold(True)

        print(self.lstColonnesCheckees)

        self.titre = QLabel("Sélection des colonnes à afficher")
        self.titre.setFont(self.font)

        self.fontbutton = QFont()
        self.fontbutton.setPointSize(12)

        self.saveButton = QPushButton("Enregistrer les modifications")
        self.saveButton.setMinimumSize(100,40)
        self.saveButton.setMaximumSize(310,100)
        self.saveButton.clicked.connect(self.save)
        self.saveButton.setFont(self.fontbutton)


        self.widgetCheckBox = QWidget()
        self.layoutCheckBox = QVBoxLayout()

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFixedSize(400,450)

        for i in range(len(self.sheetLst)):
            label = QLabel()
            nomLabel = traitementNom(self.sheetLst[i])
            contentLabel = str(self.sheetLst[i])
            setattr(self, nomLabel, label)
            getattr(self, nomLabel).setText(contentLabel)
            getattr(self, nomLabel).setFont(self.font2)

            self.layoutCheckBox.addWidget(getattr(self, nomLabel), 0, Qt.AlignmentFlag.AlignCenter)

            self.workSheet = pd.read_excel(bdd, sheet_name=self.sheetLst[i])
            self.colonnes = list(self.workSheet.columns.tolist())
            
            for j in range(len(self.colonnes)):
                checkBox = QCheckBox()
                nomCheckBox = traitementNom(self.colonnes[j])
                if nomCheckBox not in self.lstcheckcrees:
                    self.lstColonnesComplete.append(nomCheckBox)
                    contentCheckBox = str(self.colonnes[j])
                    setattr(self, nomCheckBox, checkBox)
                    getattr(self, nomCheckBox).setText(contentCheckBox)

                    if self.colonnes[j] in self.lstColonnesCheckees:
                        getattr(self, nomCheckBox).setChecked(True)

                    self.layoutCheckBox.addWidget(getattr(self,nomCheckBox), 0, Qt.AlignmentFlag.AlignLeft)
                    self.lstcheckcrees.append(nomCheckBox)
                else:
                    print(nomCheckBox)
            
            
        self.widgetCheckBox.setLayout(self.layoutCheckBox)
        self.scroll_area.setWidget(self.widgetCheckBox)

        self.layout.addWidget(self.titre, 0, Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.scroll_area, 0, Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.saveButton, 0, Qt.AlignmentFlag.AlignCenter)

        self.setLayout(self.layout)

        self.lstcheckcrees.clear()

        
    def save(self):
        self.lstSortie = []


        for i in range(len(self.lstColonnesComplete)):
            if getattr(self, self.lstColonnesComplete[i]).isChecked():
                self.lstSortie.append(getattr(self, self.lstColonnesComplete[i]).text())

        self.upload_path_parm = os.path.join(tempfile.mkdtemp(), os.path.basename(param_URL))
        with open(self.upload_path_parm, 'w') as f:
            for item in self.lstSortie:
                f.write(item + "\n")

        with open(self.upload_path_parm, 'rb') as content_file:
            file_content = content_file.read()

        file_folder = ctx.web.get_folder_by_server_relative_url("/sites/BricoDepot/Shared%20Documents/Donnees")
        target_file = file_folder.upload_file('param.txt', file_content).execute_query()

        print("File hase been uploaded to url: {0}".format(target_file.serverRelativeUrl))
		
        shutil.rmtree(os.path.dirname(self.upload_path_parm))

##### Model Pandas pour tableView -> N'accepte pas la modification #####
class pandasModel(QAbstractTableModel):
    def __init__(self, data):
        QAbstractTableModel.__init__(self) 
        self._data = data

        self.colors = dict()

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if index.isValid():
            if role == Qt.ItemDataRole.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])       
            if role == Qt.ItemDataRole.BackgroundRole:
                color = self.colors.get((index.row(), index.column()))
                if color is not None:
                    return color
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._data.columns[col]
        return None

    def change_color(self, row, column, color):
        ix = self.index(row, column)
        self.colors[(row, column)] = color
        self.dataChanged.emit(ix, ix, (Qt.ItemDataRole.BackgroundRole,))

##### Model Pandas pour tableView -> Accepte la modification #####
class pandasEditableModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, index):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]
    
    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if index.isValid():
            if role == Qt.ItemDataRole.DisplayRole or role == Qt.ItemDataRole.EditRole:
                value = self._data.iloc[index.row(), index.column()]
                return str(value)
            
    def setData(self, index, value, role):
        if role == Qt.ItemDataRole.EditRole:
            self._data.iloc[index.row(), index.column()] = value
            return True
        return False
    
    def headerData(self, col, orientation, role):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._data.columns[col]
        
    def flags(self, index):
        return Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable

if __name__ == '__main__':
    app = QApplication(sys.argv)

    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):    
        repertexec = sys._MEIPASS # programme traité par pyinstaller
    else:
        repertexec = os.path.dirname(os.path.abspath(__file__)) # non traité
    
    clean = True
    if clean==True:
        styleSheet = """QWidget
                    {
                        background-color: #464646;
                        color: #ffffff;
                        
                    }

                    /*-----QPushButton-----*/
                    QPushButton {
                                    background-color: #c2c7d5;
                                    color: #000;
                                    border: none;
                                    border-radius: 8px;
                                    padding: 3px;
                                    font-size: 20px;
                                }
                    QPushButton:hover {
                                    border: 2px solid black;
                                    color: black;
                                }
                    QPushButton:pressed {
                                    background-color: #b4c2c7d5;
                                    border: 2px solid black;
                                }


                    /*-----QCheckBox-----*/
                    QCheckBox
                    {
                        background-color: transparent;
                        color: #fff;
                        font-size: 10px;
                        font-weight: bold;
                        border: none;
                        border-radius: 5px;

                    }

                    /*-----QLineEdit-----*/
                    QLineEdit
                    {
                        background-color: #c2c7d5;
                        color: #000;
                        border: none;
                        border-radius: 8px;
                        padding: 3px;
                    }


                    /*-----QListView-----*/
                    QListView
                    {
                        background-color: #333333;
                        color: #fff;
                        font-size: 12px;
                        font-weight: bold;
                        border: 1px solid #191919;
                        show-decoration-selected: 0;
                        padding-left: -13px;
                        padding-right: -13px;

                    }


                    QListView::item
                    {
                        color: #888b8b;
                        background-color: #454e5e;
                        border: none;
                        padding: 5px;
                        border-radius: 0px;
                        padding-left : 10px;
                        height: 42px;

                    }

                    QListView::item:selected
                    {
                        color: #000;
                        background-color: #8a8a8a;

                    }


                    QListView::item:!selected
                    {
                        color:white;
                        background-color: transparent;
                        border: none;
                        padding-left : 10px;

                    }


                    QListView::item:!selected:hover
                    {
                        color: #000;
                        background-color: #bcbdbb;
                        border: none;
                        padding-left : 10px;

                    }

                    /*-----QTableView & QTableWidget-----*/
                    QTableView
                    {
                        background-color: #252525;
                        border: 1px solid gray;
                        color: #f0f0f0;
                        gridline-color: gray;
                        outline : 0;

                    }


                    QTableView::disabled
                    {
                        background-color: #242526;
                        border: 1px solid #32414B;
                        color: #656565;
                        gridline-color: #656565;
                        outline : 0;

                    }


                    QTableView::item:hover 
                    {
                        background-color: #bcbdbb;
                        color: #000;

                    }


                    QTableView::item:selected 
                    {
                        background-color: #c2c7d5;
                        color: #000;

                    }


                    QTableView::item:selected:disabled
                    {
                        background-color: #1a1b1c;
                        border: 2px solid #525251;
                        color: #656565;

                    }


                    QTableCornerButton::section
                    {
                        background-color: #343a49;
                        color: #fff;

                    }


                    QHeaderView::section
                    {
                        color: #fff;
                        border-top: 0px;
                        border-bottom: 1px solid gray;
                        border-right: 1px solid gray;
                        background-color: #343a49;
                        margin-top:1px;
                        margin-bottom:1px;
                        padding: 5px;

                    }


                    QHeaderView::section:disabled
                    {
                        background-color: #525251;
                        color: #656565;

                    }


                    QHeaderView::section:checked
                    {
                        color: #000;
                        background-color: #b6b0b0;

                    }


                    QHeaderView::section:checked:disabled
                    {
                        color: #656565;
                        background-color: #525251;

                    }


                    QHeaderView::section::vertical::first,
                    QHeaderView::section::vertical::only-one
                    {
                        border-top: 1px solid #353635;

                    }


                    QHeaderView::section::vertical
                    {
                        border-top: 1px solid #353635;

                    }


                    QHeaderView::section::horizontal::first,
                    QHeaderView::section::horizontal::only-one
                    {
                        border-left: 1px solid #353635;

                    }


                    QHeaderView::section::horizontal
                    {
                        border-left: 1px solid #353635;

                    }


                    /*-----QScrollBar-----*/
                    QScrollBar:horizontal 
                    {
                        background-color: transparent;
                        height: 14px;
                        margin: 0px;
                        padding: 0px;

                    }


                    QScrollBar::handle:horizontal 
                    {
                        border: none;
                        min-width: 100px;
                        background-color: #9b9b9b;

                    }


                    QScrollBar::add-line:horizontal, 
                    QScrollBar::sub-line:horizontal,
                    QScrollBar::add-page:horizontal, 
                    QScrollBar::sub-page:horizontal 
                    {
                        width: 0px;
                        background-color: transparent;

                    }


                    QScrollBar:vertical 
                    {
                        background-color: transparent;
                        width: 14px;
                        margin: 0;

                    }


                    QScrollBar::handle:vertical 
                    {
                        border: none;
                        min-height: 100px;
                        background-color: #9b9b9b;

                    }


                    QScrollBar::add-line:vertical, 
                    QScrollBar::sub-line:vertical,
                    QScrollBar::add-page:vertical, 
                    QScrollBar::sub-page:vertical 
                    {
                        height: 0px;
                        background-color: transparent;

                    }
                    """

    splash = SplashScreen()
    splash.show()

    app.processEvents()
    splash.labelDescription.setText('<strong>Connexion à Sharepoint</strong>')

    #####ETAPE 1: Création context authentification#####
    username = "acae250d-01e9-4f32-9d65-e06fa388ff60"
    password = "8FG7d+Es/DYXCJWN8spbNV6qyU5TQqUsoKmg5HLsHw4="
    test_team_site_url = "https://sgzkl.sharepoint.com/sites/BricoDepot"
    bdd_URL = "/sites/BricoDepot/Shared%20Documents/Donnees/BDD.xlsx"
    mdp_URL = "/sites/BricoDepot/Shared%20Documents/Donnees/MDP.xlsx"
    requete_URL = "/sites/BricoDepot/Shared%20Documents/Donnees/REQ.xlsx"
    param_URL = "/sites/BricoDepot/Shared%20Documents/Donnees/param.txt"

    ctx = ClientContext(test_team_site_url).with_credentials(ClientCredential(username, password))
    web = ctx.web
    ctx.load(web).execute_query()
    print ("Connexion à Sharepoint réussie")

    splash.progressBar.setValue(10)
    splash.labelDescription.setText('<strong>Récupération fichier 1</strong>')

    #####ETAPE 2: Récupération BDD#####
    responseBDD = File.open_binary(ctx, bdd_URL)
    print("Reponse trouvée")
    bytes_file_obj_bdd = io.BytesIO()
    bytes_file_obj_bdd.write(responseBDD.content)
    bytes_file_obj_bdd.seek(0)
    print ("BDD chargée")

    splash.progressBar.setValue(20)
    splash.labelDescription.setText('<strong>Récupération fichier 2</strong>')
    
    #####ETAPE 3: Récupération MDP#####
    responseMDP = File.open_binary(ctx, mdp_URL)
    print("Reponse trouvée")
    bytes_file_obj_mdp = io.BytesIO()
    bytes_file_obj_mdp.write(responseMDP.content)
    bytes_file_obj_mdp.seek(0)
    print ("MDP chargés")

    splash.progressBar.setValue(30)
    splash.labelDescription.setText('<strong>Récupération fichier 3</strong>')

    #####ETAPE 4: Récupération requetes#####
    responseREQ = File.open_binary(ctx, requete_URL)
    print("Reponse trouvée")
    bytes_file_obj_req = io.BytesIO()
    bytes_file_obj_req.write(responseREQ.content)
    bytes_file_obj_req.seek(0)
    print ("Requêtes chargées")

    bdd = bytes_file_obj_bdd
    mdp = bytes_file_obj_mdp
    req = bytes_file_obj_req

    splash.progressBar.setValue(40)
    splash.labelDescription.setText('<strong>Fin de paramétrage</strong>')

    #####ETAPE 5: Création de la dataframe#####
    xlsx_file = pd.ExcelFile(bdd)
    sheet_lst = xlsx_file.sheet_names
    #splash.setProgress(50)

    sheet_Globale = pd.DataFrame()
    ##TODO: vérif
    bool_premier = True
    sheet_lst = [element for element in sheet_lst if element not in ['Accueil', 'BDD']]
    increment = 60 / (len(sheet_lst)+1) 
    print(len(sheet_lst))

    for i in range(len(sheet_lst)):
            app.processEvents()
            splash.labelSubDescription.setText('<strong>' + sheet_lst[i] + ' ' + str(i+1) + '/' + str(len(sheet_lst)) + '</strong>')
            splash.progressBar.setValue(round(i*increment) + 40)
            
            workSheet = pd.read_excel(bdd, sheet_name=sheet_lst[i])
            common_col = ["Code BRICO", "Code EASIER", "Dépôt", "Région 2022"]
            if bool_premier == True:
                sheet_Globale = workSheet
                bool_premier=False
            else:
                sheet_Globale = pd.merge(sheet_Globale, workSheet, on=common_col)

    splash.labelSubDescription.setText('')
    splash.progressBar.setValue(100)
    time.sleep(1)
    
    def traitementNom(input):
        trait = unidecode(input)
        trait = trait.replace(' ','')
        trait = trait.replace("'", "")
        trait = trait.replace("(","")
        trait = trait.replace(")","")
        trait = trait.replace("/", "")
        trait = trait.replace("-", "")
        return trait
    
    def resource_path(relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    app.setStyleSheet(styleSheet)

    #####ETAPE 6: Récupération des paramètre#####
    downParam = os.path.join(tempfile.mkdtemp(), os.path.basename(param_URL))
    with open(downParam, "wb") as local_file:
        file = ctx.web.get_file_by_server_relative_url(param_URL).download(local_file).execute_query()
    print("[Ok] file has been downloaded into: {0}".format(downParam))
    
    ##Lecture des paramètres
    with open(downParam, 'r') as f:
        contenu = f.read()
        
    listeParam = contenu.split('\n')[:-1]
    shutil.rmtree(os.path.dirname(downParam))

    splash.close()
    main = mainWindow()
    main.lstParam = listeParam
    form = logWindow()
    form.show()
    

    sys.exit(app.exec())  