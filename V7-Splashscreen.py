import pandas as pd
import openpyxl as wb
import sys, io, shutil, tempfile, os, datetime, time
from unidecode import unidecode

from PyQt6.QtCore import Qt, QSize, QSortFilterProxyModel, QAbstractTableModel, QTimer, QThread, pyqtSignal
from PyQt6.QtGui import QIcon, QPainter, QColor, QPixmap, QBrush, QFont
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QProgressBar, QHeaderView, QMessageBox, QHBoxLayout, QWidget, QLineEdit, QGridLayout, QComboBox, QVBoxLayout, QStackedWidget, QScrollArea, QFrame, QTableView, QSpacerItem 
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File 

class logWindow(QWidget):  
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Authentification")
        self.setFixedSize(570,200)
        self.w = None 
        self.role = ""

        ########## Gathering the account DataBase ##########
        self.ddbmdp = pd.read_excel(mdp, sheet_name='MDPFINAUX')

        ########## logV1 ##########
        self.logv1 = QWidget()
        
        self.hLayout1 = QHBoxLayout()

        self.logoLabel1 = QLabel("")
        self.logoLabel1.setMaximumSize(QSize(200,140))
        self.logoLabel1.setPixmap(QPixmap("logo.png"))
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


        ########## logV2 ##########
        self.logv2 = QWidget()
        
        self.hLayout2 = QHBoxLayout()

        self.logoLabel2 = QLabel()
        self.logoLabel2.setMaximumSize(QSize(200,140))
        self.logoLabel2.setPixmap(QPixmap("logo.png"))
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
                    self.Stack.setCurrentIndex(2)
                    if self.w is None:
                        self.w=main
                        main.user = self.inputUser.text()
                        main.password = self.inputPassword.text()
                        main.role = role
                        main.denom = denom
                        main.demarrage()
                        #self.w=mainWindow()
                        #self.w.user = self.inputUser.text()
                        #self.w.password = self.inputPassword.text()
                        #self.w.role = role
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

class mainWindow(QWidget):

    def __init__(self):
        super().__init__()

        self.setFixedSize(720,440)

        ##### Important variables #####
        self.user = ""
        self.password = ""
        self.role = ""
        self.denom = ""

        ##### Loading and filtering the data #####
        self.loadExcel()

        ########## STYLESHEET ##########
        #self.setStyleSheet("""
            #QLineEdit{
            #    font-size: 20px
            #}
            #QPushButton{
            #    font-size: 20px
            #}
            #""")

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
        ##### Loading #####
        excel = pd.ExcelFile(bdd)
        lst_sheet = excel.sheet_names
        self.d = {}
        for i in range(len(lst_sheet)):
            if lst_sheet[i] != "Accueil" and lst_sheet[i] != "BDD":
                self.d[lst_sheet[i]] = pd.read_excel(bdd, sheet_name=lst_sheet[i])

        ##### Gathering the number of request for the admin interface #####
        ##TODO: Récupérer le nombre de requete, VERIFIER SI C'EST BON
        #self.df = pd.read_excel(req, sheet_name='Requete')
        #self.lstReq = self.df[1].values.tolist()
        #self.nbrReq = len(self.lstReq)

        ##### Generation of the dataFrame #####
        self.loadExcel()

        ##### Generation of the layout #####
        self.loadGUI()
        
    ##### Function used to load the Excel data file #####
    def loadExcel(self):
        self.sheet = sheet_Globale
        self.sheet_columns = self.sheet.columns.to_list()

        if self.role != "Admin":
            if self.role == "Région":
                ##### Gathering only the lines for a specific regional manager
                self.sheet_tri = self.sheet.loc[self.sheet['Directeur Régional'] == self.denom]
                print (self.sheet_tri)
                self.model = pandasModel(self.sheet_tri)
            elif self.role == "Dépôt":
                ##### Gathering only the lines for a site #####
                self.sheet_tri = self.sheet.loc[self.sheet['Directeur dépôt'] == self.denom]
                self.model = pandasModel(self.sheet_tri)
        else:
            #self.model = pandasModel(self.sheet)
            self.model = pandasEditableModel(self.sheet)
            print("Model créé Admin en lecture écriture")

    ##### Function used to create the interface dynamically #####
    def loadGUI(self):
        
        if self.role == "Admin":
            print("Creation du GUI ADMIN")
            ##### Creation of the admin interface #####

            self.adminGUI = QWidget()

            ##TODO: Corriger le titre qui apparait pas en entier
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
            self.logo.setPixmap(QPixmap("logo.png"))
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
            #self.topBanner.setStyleSheet("background-color: red")

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
            self.proxy_model.setFilterKeyColumn(-1) #Toutes les colonnes
            self.proxy_model.setSourceModel(self.model)
            self.proxy_model.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

            self.tab.setModel(self.proxy_model)
            self.tab.resizeColumnsToContents()

            self.searchBar.textChanged.connect(self.proxy_model.setFilterFixedString)

            self.middleLayout.addWidget(self.searchBar)
            self.middle.setLayout(self.middleLayout)
            #self.middle.setStyleSheet("background-color: green;")

            ##### Push buttons #####
            fontBouton = QFont()
            fontBouton.setPointSize(12)
            fontBouton.setBold(True)
            fontBouton.setWeight(500)

            self.creer = QPushButton("Créer un dépôt")
            self.creer.clicked.connect(self.creerDepot)
            self.creer.setMaximumSize(QSize(300, 100))
            self.creer.setMinimumSize(QSize(300,40))
            #self.creer.setStyleSheet(styleSheetBouton)
            iconCreer = QIcon()
            iconCreer.addPixmap(QPixmap("Icons/create.png"), QIcon.Mode.Normal, QIcon.State.Off)
            self.creer.setIcon(iconCreer)

            self.modifier = QPushButton("Modifier un dépôt")
            self.modifier.setFont(fontBouton)
            self.modifier.clicked.connect(self.modifDepot)
            self.modifier.setMaximumSize(QSize(300, 100))
            #self.modifier.setStyleSheet(styleSheetBouton)
            iconModifier = QIcon()
            iconModifier.addPixmap(QPixmap("Icons/edit.png"), QIcon.Mode.Normal, QIcon.State.Off)
            self.modifier.setIcon(iconModifier)

            self.supprimer = QPushButton("Supprimer un dépôt")
            self.supprimer.clicked.connect(self.supprimerDepot)
            self.supprimer.setMaximumSize(QSize(300, 100))
            #self.supprimer.setStyleSheet(styleSheetBouton)
            iconSupprimer = QIcon()
            iconSupprimer.addPixmap(QPixmap("Icons/delete.png"), QIcon.Mode.Normal, QIcon.State.Off)
            self.supprimer.setIcon(iconSupprimer)

            self.bandeau = QWidget()
            self.bandeauBoutons = QHBoxLayout()
            self.bandeauBoutons.addWidget(self.creer)
            self.bandeauBoutons.addWidget(self.modifier)
            self.bandeauBoutons.addWidget(self.supprimer)

            self.bandeau.setLayout(self.bandeauBoutons)
            #self.bandeau.setStyleSheet("background-color: blue;")

            self.boutonValidation = QPushButton("Confirmer modifications")
            self.boutonValidation.setMaximumSize(QSize(300,100))
            self.boutonValidation.setMinimumSize(QSize(300,40))
            #self.boutonValidation.setStyleSheet(styleSheetBouton)
            self.boutonValidation.clicked.connect(self.saveData)
            iconValider = QIcon()
            iconValider.addPixmap(QPixmap("Icons/check.png"), QIcon.Mode.Normal, QIcon.State.Off)
            self.boutonValidation.setIcon(iconValider)

            ##Comptage nombre de requêtes de changement
            self.sheetRequetes = pd.read_excel(req, sheet_name="Requete")
            self.nbrLignes = self.sheetRequetes.shape[0]

            self.traitementRequetes = QPushButton("Requêtes de MAJ: " + str(self.nbrLignes))
            self.traitementRequetes.setMaximumSize(QSize(300,100))
            #self.traitementRequetes.setStyleSheet(styleSheetBouton)
            ##TODO: Rajouter une bulle dans le texte pour ca
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
            #self.bandeauInf.setStyleSheet("background-color: purple;")
            

            ##### Setting up the Widget #####
            self.layoutAdminGUI = QVBoxLayout()
            self.layoutAdminGUI.addWidget(self.topBanner, 0, Qt.AlignmentFlag.AlignHCenter)
            self.layoutAdminGUI.addWidget(self.middle)
            self.layoutAdminGUI.addWidget(self.tab)
            self.layoutAdminGUI.addWidget(self.bandeau)
            self.layoutAdminGUI.addWidget(self.bandeauInf)

            self.layoutAdminGUI.setSpacing(0)
            self.adminGUI.setLayout(self.layoutAdminGUI)

            ##### Adding the interface to the StackedWidget #####

            self.Stack.addWidget(self.adminGUI)
            self.Stack.setCurrentIndex(1)

        else:
            print("Creation du GUI en lecture")
            ##### Creation of the regular interface #####
            self.adminGUI = QWidget()

            ##TODO: Corriger le titre qui apparait pas en entier
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
            self.logo.setPixmap(QPixmap("logo.png"))
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
            #self.requete.setStyleSheet(styleSheetBouton)
            self.requete.clicked.connect(self.demandeChangement)

            self.bandeau = QWidget()
            self.bandeauBoutons = QHBoxLayout()
            self.bandeauBoutons.addWidget(self.requete)
            self.bandeau.setLayout(self.bandeauBoutons)

            ##### Setting up the Widget #####
            self.layoutAdminGUI = QVBoxLayout()
            self.layoutAdminGUI.addWidget(self.topBanner, 0, Qt.AlignmentFlag.AlignHCenter)
            self.layoutAdminGUI.addWidget(self.middle)
            self.layoutAdminGUI.addWidget(self.bandeau)
            self.adminGUI.setLayout(self.layoutAdminGUI)

            ##### Adding the interface to the StackedWidget #####
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
        modifWindow.sheet=self.sheet
        self.w = modifWindow()
        self.w.show()

    def supprimerDepot(self):
        print("Suppression")
        suppWindow.sheet=self.sheet
        self.w = suppWindow()
        self.w.show()

    def demandeChangement(self):
        print("Formulaire demande changement")
        demandeChangement.sheet = self.sheet_tri
        self.w = demandeChangement()
        self.w.show()

    def traitementRequetesChangement(self):
        print("Traitement des requetes de changement")
        self.w = traitementDemandeChangement()
        self.w.show()

    def chargerModif(self):
        self.model = pandasEditableModel(self.sheet)
        self.proxy_model.setSourceModel(self.model)
        self.tab.setModel(self.proxy_model)
        self.tab.resizeColumnsToContents()
        print("Model chargé")

    def saveData(self):
        print("Saving...")
        print(self.sheet)
        
        #Storage du fichier excel temporaire, changer pour le storer uniquement si c'est un utilisateur admin
        self.download_path = os.path.join(tempfile.mkdtemp(), os.path.basename(bdd_URL))
        with open(self.download_path, "wb") as local_file:
            file = ctx.web.get_file_by_server_relative_url(bdd_URL).download(local_file).execute_query()
        print("[Ok] file has been downloaded into: {0}".format(self.download_path))

        self.wb_obj = wb.load_workbook(self.download_path)
        print("Workbook loadé sur openpyxl")

        my_set = set(self.sheet_columns)

        for key in self.d.keys():
            lst_columns = self.d[key].columns.to_list()
            lst_clean = []
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

        print("nombre row: " + str(nbRow) + " et le reste: " + str(reste))
        print("nombre max index : " + str(len(self.sheets)))

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
                #print("sheet_name: " + self.sheet_name)
                #print(f"Index = {index}, sheet_name = {self.sheet_name}")
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
        ##TODO: corriger le code brico et tout
        #Si la fenêtre existe déjà
        if str in self.windowslst:
            self.w = self.windowslst[str]
        else:
            fenetre = testDeLectureCreation()
            nomFenetre = str
            setattr(self, nomFenetre, fenetre)
            getattr(self,nomFenetre).excel_sheet = str
            getattr(self,nomFenetre).demarrage(str)
            self.windowslst[str] = getattr(self,nomFenetre)
            self.w = getattr(self, nomFenetre)

        if crea == False:
            self.w.show()


    def confirmerCreation(self):
        print("ca marche")
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
        print (CreaSheet)

        main.sheet = CreaSheet
        main.chargerModif()
        self.close()
        

    def closeEvent(self, event):
        print("je fais le close event")
        self.w=''
        if self.w:
            self.w.close()


class testDeLectureCreation(QWidget):
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
            #faire la scroll area
            print("Scroll Area creation")
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

class modifWindow(QWidget):
    sheet = pd.DataFrame

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Modification d'un dépôt")
        self.title = QLabel("Modification d'un dépôt")
        self.setFixedSize(300,80)

        ##TODO: Faire un stacked widget
        self.stack = QStackedWidget()
                
        ##TODO: Widget 1 -> Sélection du dépot
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
        
        ##TODO: Widget 2 -> Affichage du dépot + possibilité de toujours le modifier
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
            self.sheet[listHeaders[i]][rowIndex] = self.sheet_tri[listHeaders[i]][rowIndex]
                
    ##Transmission de la nouvelle sheet à la fenêtre principale
        main.sheet = self.sheet
        main.chargerModif()

        self.close()

    def center(self):
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()

        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def annuler(self):
        self.close()
    
##TODO: Performance enhancement
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
        ##TODO: Widget 1 -> Sélection du dépot
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
        
        ##TODO: Widget 2 -> Affichage du dépot + possibilité de toujours le modifier
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

##TODO: UI Design + Rajouter icones dans les boutons
class demandeChangement(QWidget):
    sheet = pd.DataFrame
    def __init__(self):
        super().__init__()
        self.setFixedSize(1280,720)
        

        self.setWindowTitle("Demande de changement de données")

        self.layout = QVBoxLayout()

        self.Titre = QLabel("Demande de changement de données")
        font = QFont()
        font.setPointSize(28)
        font.setBold(True)
        font.setWeight(75)
        self.Titre.setFont(font)

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
        #self.valider.setStyleSheet(styleSheetBouton)
        self.valider.clicked.connect(self.transmettre)

        self.annuler = QPushButton("Annuler la demande de\nchangement de données")
        self.annuler.setMaximumSize(400,100)
        #self.annuler.setStyleSheet(styleSheetBouton)
        self.annuler.clicked.connect(self.annulerF)

        self.space = QSpacerItem(100,0)

        self.bandeauBoutonsLayout.addItem(self.space)
        self.bandeauBoutonsLayout.addWidget(self.annuler)
        self.bandeauBoutonsLayout.addWidget(self.valider)
        self.bandeauBoutonsLayout.addItem(self.space)

        self.bandeauBouton.setLayout(self.bandeauBoutonsLayout)
        
        self.layout.addWidget(self.Titre, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout.addWidget(self.texte, 0, Qt.AlignmentFlag.AlignHCenter)
        self.layout.addWidget(self.table)
        self.layout.addWidget(self.bandeauBouton)
        

        self.layout.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)

        self.setLayout(self.layout)

    def transmettre(self):

        self.dataframereq = pd.read_excel(req, sheet_name='Requete')
        print(self.dataframereq)

        self.sheet['Utilisateur'] = main.denom
        self.sheet['Date demande'] = datetime.now().strftime('%Y-%m-%d')

        self.newDF = pd.DataFrame
        self.newDF = pd.concat([self.dataframereq,self.sheet], axis=0)
        self.newDF = self.newDF.reset_index(drop=True)
        print(self.newDF)
        
        ##TODO: Faire un comparatif avec la dataframe triée de base. Pour chaque ligne checker si il ya eu un changement. Si 1 changement detecté inserer direct. 

        #Upload sur sharepoint
        download_path_req = os.path.join(tempfile.mkdtemp(), os.path.basename(bdd_URL))
        with open(download_path_req, "wb") as local_file:
            file = ctx.web.get_file_by_server_relative_url(requete_URL).download(local_file).execute_query()
        print("[Ok] file has been downloaded into: {0}".format(download_path_req))

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

    def annulerF(self):
        self.close()

class traitementDemandeChangement(QWidget):

    def __init__(self):
        super().__init__()

        ##### Gathering the requests #####
        self.sheet = main.sheet
        self.setWindowTitle("Traitement demandes de MAJ")

        self.dfRequete = pd.read_excel(req, sheet_name='Requete')
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
        #self.boutonValidation.setStyleSheet(styleSheetBouton)

        self.boutonRetirer = QPushButton("Retirer la requête")
        self.boutonRetirer.clicked.connect(self.retirer)
        self.boutonRetirer.setMaximumSize(300,100)
        self.boutonRetirer.setMinimumWidth(200)
        #self.boutonRetirer.setStyleSheet(styleSheetBouton)

        self.boutonAnnuler = QPushButton("Annuler")
        self.boutonAnnuler.clicked.connect(self.annuler)
        self.boutonAnnuler.setMaximumSize(300,100)
        self.boutonAnnuler.setMinimumWidth(200)
        #self.boutonAnnuler.setStyleSheet(styleSheetBouton)

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
        ##TODO: ca charge deu fois la fonction faut retirer
        ##TODO: Essayer de freeze la premiere colonne
        ##TODO: Surligner les parties qui changent
        #Création de la sheet
        depot = self.comboBox2.currentText()
        print("le depot: " + depot)
        depot = str(depot.split('-')[1])

        #Ligne actuelle
        self.actualdata = self.sheet.loc[self.sheet['Dépôt'] == depot]
        self.actualdata.insert(0, 'TypeLigne', ['Données actuelles'])
        print(self.actualdata)

        #Ligne de changment
        self.sheetRequete = self.dfRequete.loc[self.dfRequete['Dépôt'] == depot]
        #self.sheetRequete.insert(0, 'TypeLigne', ['Requête de changement'])


        self.sheet_tri = pd.concat([self.actualdata, self.sheetRequete], axis=0)
        self.sheet_tri = self.sheet_tri.reset_index(drop=True)
        print (self.sheet_tri)
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
        #Dans ce cas on va tout simplement remplacer dans la sheet de base du main
        lst_colonnes = main.sheet.columns.tolist()
        self.codeBrico = self.sheet_tri.iloc[1]["Code BRICO"]
        self.depot = self.sheet_tri.iloc[1]["Dépôt"]
        self.utilisateur = self.sheet_tri.iloc[1]["Utilisateur"]
        self.dateDemande = self.sheet_tri.iloc[1]["Date demande"]

        #Index de la ligne à changer
        index_ligne = self.dfRequete.loc[(self.dfRequete['Code BRICO'] == self.codeBrico) & (self.dfRequete['Utilisateur'] == self.utilisateur) & (self.dfRequete['Date demande'] == self.dateDemande)].index[0]
        indexMain = main.sheet.loc[(main.sheet['Code BRICO'] == self.codeBrico)].index[0]

        for i in range(len(lst_colonnes)):
            main.sheet[lst_colonnes[i]][indexMain] = self.dfRequete[lst_colonnes[i]][index_ligne]

        main.chargerModif()

        #La requête est passée, supprimer la requete et supprimer tout bref
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
        self.comboBox2.setCurrentIndex(curInd-1)

    ##TODO: Refaire le test avec 1 requete et plusieurs requetes
    
    def retirer(self):
        print("retirer la requete")
        self.codeBrico = self.sheet_tri.iloc[1]["Code BRICO"]
        self.depot = self.sheet_tri.iloc[1]["Dépôt"]
        self.utilisateur = self.sheet_tri.iloc[1]["Utilisateur"]
        self.dateDemande = self.sheet_tri.iloc[1]["Date demande"]

        print("Code brico: " + str(self.codeBrico))
        print("Depot: " + str(self.depot))
        print("Utilisateur: " + str(self.utilisateur))
        print("Date demande: " + self.dateDemande)

        index_ligne = self.dfRequete.loc[(self.dfRequete['Code BRICO'] == self.codeBrico) & (self.dfRequete['Utilisateur'] == self.utilisateur) & (self.dfRequete['Date demande'] == self.dateDemande)].index[0]
        print("index de la ligne: "+ str(index_ligne))
        self.dfRequete = self.dfRequete.drop(index_ligne)
        self.dfRequete.reset_index(drop=True)
        print(self.dfRequete)


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

        if self.dfRequete.empty:
            QMessageBox.information(self, 'Succès', 'Plus de requêtes à traiter, fermeture...')
            self.close()
            return
        
        curInd = self.comboBox2.currentIndex()
        self.comboBox2.removeItem(curInd)
        self.comboBox2.setCurrentIndex(curInd-1)

    def center(self):
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()

        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def annuler(self):
        self.close()   



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

    username = "acae250d-01e9-4f32-9d65-e06fa388ff60"
    password = "8FG7d+Es/DYXCJWN8spbNV6qyU5TQqUsoKmg5HLsHw4="
    test_team_site_url = "https://sgzkl.sharepoint.com/sites/BricoDepot"
    bdd_URL = "/sites/BricoDepot/Shared%20Documents/Donnees/BDD.xlsx"
    mdp_URL = "/sites/BricoDepot/Shared%20Documents/Donnees/MDP.xlsx"
    requete_URL = "/sites/BricoDepot/Shared%20Documents/Donnees/REQ.xlsx"

    ctx = ClientContext(test_team_site_url).with_credentials(ClientCredential(username, password))
    web = ctx.web
    ctx.load(web).execute_query()
    print ("Connexion à Sharepoint réussie")

    #splash.setProgress(10)

    responseBDD = File.open_binary(ctx, bdd_URL)
    print("Reponse trouvée")
    bytes_file_obj_bdd = io.BytesIO()
    bytes_file_obj_bdd.write(responseBDD.content)
    bytes_file_obj_bdd.seek(0)
    print("BDD chargée")
    #splash.setProgress(20)

    responseMDP = File.open_binary(ctx, mdp_URL)
    print("Reponse trouvée")
    bytes_file_obj_mdp = io.BytesIO()
    bytes_file_obj_mdp.write(responseMDP.content)
    bytes_file_obj_mdp.seek(0)
    print ("MDP chargés")
    #splash.setProgress(30)

    responseREQ = File.open_binary(ctx, requete_URL)
    print("Reponse trouvée")
    bytes_file_obj_req = io.BytesIO()
    bytes_file_obj_req.write(responseREQ.content)
    bytes_file_obj_req.seek(0)
    print ("Requêtes chargées")
    #splash.setProgress(40)

    ##### Storing the data streams #####
    bdd = bytes_file_obj_bdd
    mdp = bytes_file_obj_mdp
    req = bytes_file_obj_req

    print("Lecture stream BDD pour création du set de données")
    xlsx_file = pd.ExcelFile(bdd)
    sheet_lst = xlsx_file.sheet_names
    #splash.setProgress(50)

    sheet_Globale = pd.DataFrame()
    ##TODO: vérif
    bool_premier = True
    for i in range(len(sheet_lst)):
        if sheet_lst[i] != 'Accueil' and sheet_lst[i] != 'BDD':
            print (sheet_lst[i])
            workSheet = pd.read_excel(bdd, sheet_name=sheet_lst[i])
            common_col = ["Code BRICO", "Code EASIER", "Dépôt", "Région 2022"]
            if bool_premier == True:
                sheet_Globale = workSheet
                bool_premier=False
            else:
                sheet_Globale = pd.merge(sheet_Globale, workSheet, on=common_col)
    
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

    main = mainWindow()
    form = logWindow()
    form.show()
    

    sys.exit(app.exec())




    
