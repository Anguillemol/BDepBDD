import sys
from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from PyQt6.uic import loadUi

import pandas as pd
import numpy as np
import openpyxl

class logWindow(QWidget):
    

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Authentification BricoDepot")
        self.w = None 
        self.role = ""
        self.submitClicked = pyqtSignal(str,str,str)

        ########## Gathering the account DataBase ##########
        self.ddbmdp = pd.read_excel('C:/Users/lucas/Test.xlsx', sheet_name='MDP')      
        
        ##### Static text #####
        title = QLabel("Authentification BDD")
        user = QLabel("Nom d'utilisateur:")
        password = QLabel("Mot de passe:")
        title2 = QLabel("Authentification BDD")
        user2 = QLabel("Nom d'utilisateur:")
        password2 = QLabel("Mot de passe:")
        error = QLabel("Informations invalides !")
        error.setStyleSheet("color: red; font-weight: bold")
        connexionReussie = QLabel("Connexion établie")
        connexionReussie.setStyleSheet("font-size: 20px")

        ##### Input areas #####
        self.inputUser = QLineEdit()
        self.inputPassword = QLineEdit()
        self.inputUser2 = QLineEdit()
        self.inputPassword2 = QLineEdit()
        self.inputPassword.setEchoMode(QLineEdit.EchoMode.Password)
        self.inputPassword2.setEchoMode(QLineEdit.EchoMode.Password)

        ##### Buttons #####
        buttonCo = QPushButton("Connexion")
        buttonCo.setObjectName("button1")
        buttonCo.clicked.connect(self.PushCo)
        buttonCl = QPushButton("Nettoyer")
        buttonCl.clicked.connect(self.PushCl)

        buttonCo2 = QPushButton("Connexion")
        buttonCo2.setObjectName("button2")
        buttonCo2.clicked.connect(self.PushCo)
        buttonCl2 = QPushButton("Nettoyer")
        buttonCl2.clicked.connect(self.PushCl)

        ########## logV1 ##########
        self.logv1 = QWidget()

        layoutv1 = QGridLayout()
        layoutv1.setSpacing(3)
        
        layoutv1.addWidget(title, 0, 0, 1, 3, Qt.AlignmentFlag.AlignCenter)
        layoutv1.addWidget(user, 1, 0)
        layoutv1.addWidget(password, 2, 0)
        layoutv1.addWidget(self.inputUser, 1, 1, 1, 2)
        layoutv1.addWidget(self.inputPassword, 2, 1, 1, 2)
        layoutv1.addWidget(buttonCo, 3, 2)
        layoutv1.addWidget(buttonCl, 3, 1)

        self.logv1.setLayout(layoutv1)

        ########## logV2 ##########
        self.logv2 = QWidget()

        layoutv2 = QGridLayout()
        layoutv2.setSpacing(3)

        layoutv2.addWidget(title2, 0, 0, 1, 3, Qt.AlignmentFlag.AlignCenter)
        layoutv2.addWidget(user2, 1, 0)
        layoutv2.addWidget(password2, 2, 0)
        layoutv2.addWidget(self.inputUser2, 1, 1, 1, 2)
        layoutv2.addWidget(self.inputPassword2, 2, 1, 1, 2)
        layoutv2.addWidget(error, 3, 1, 1, 2)
        layoutv2.addWidget(buttonCo2, 4, 2)
        layoutv2.addWidget(buttonCl2, 4, 1)

        self.logv2.setLayout(layoutv2)

        ########## logFrozen ##########
        self.logFrozen = QWidget()

        layoutvFrozen = QGridLayout()
        layoutvFrozen.addWidget(connexionReussie, 0, 0, 3, 3, Qt.AlignmentFlag.AlignCenter)

        self.logFrozen.setLayout(layoutvFrozen)

        ############## StackedWidget ##############
        self.Stack = QStackedWidget (self)
        self.Stack.addWidget(self.logv1)
        self.Stack.addWidget(self.logv2)
        self.Stack.addWidget(self.logFrozen)

        self.Stack.setCurrentIndex(0)

        layoutHome = QGridLayout()
        layoutHome.addWidget(self.Stack, 0, 0)

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
                if self.inputPassword.text() == self.ddbmdp.loc[self.ddbmdp['Username'] == self.inputUser.text()]['Password'][nrow]:
                    self.Stack.setCurrentIndex(2)
                    if self.w is None:
                        self.w=mainWindow()
                        self.w.user = self.inputUser.text()
                        self.w.password = self.inputPassword.text()
                        self.w.role = role
                    self.w.show()
                    self.close()
                else:
                    self.Stack.setCurrentIndex(1)
        if name=="button2":
            if not (self.ddbmdp.loc[self.ddbmdp['Username'] == self.inputUser2.text()].empty):
                nrow = self.ddbmdp[self.ddbmdp['Username'] == self.inputUser2.text()].index.values.astype(int)[0]
                role = self.ddbmdp['Role'][nrow]
                if self.inputPassword2.text() == self.ddbmdp.loc[self.ddbmdp['Username'] == self.inputUser2.text()]['Password'][nrow]:
                    self.Stack.setCurrentIndex(2)
                    if self.w is None:
                        self.w=mainWindow()
                        self.w.user = self.inputUser2.text()
                        self.w.password = self.inputPassword2.text()
                        self.w.role = role
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

##TODO: Retirer l'écran annexe pour le chargement des données <!!>
class mainWindow(QWidget):
    def __init__(self):
        self.submitClicked = pyqtSignal(str,str,str)
        super().__init__()

        self.setFixedSize(720,440)

        ##### Important variables #####
        self.user = ""
        self.password = ""
        self.role = ""

        ##### Loading and filtering the data #####
        self.loadExcel()

        ########## STYLESHEET ##########
        self.setStyleSheet("""
            QLineEdit{
                font-size: 30px
            }
            QPushButton{
                font-size: 30px
            }
            """)

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
        
        self.Stack.setCurrentIndex(0)
        self.mainLayout.addWidget(self.Stack)

        self.setLayout(self.mainLayout)
        

        
    def demarrage(self):
        print(self.user)
        print(self.password)
        print(self.role)
        self.setFixedSize(1280,720)
        self.center()
        #TODO: Recentrer la fenêtre au milieu de l'écran
        
        ##### Generation of the dataFrame #####
        self.loadExcel()

        ##### Generation of the layout #####
        self.loadGUI()
        
    ##### Function used to load the Excel data file #####
    def loadExcel(self):
        self.sheet = pd.read_excel('C:/Solutec/1/BDD.xlsx', sheet_name='BDD')

        if self.role != "Admin":
            if self.role == "Region":
                ##### Gathering only the lines for a specific regional manager
                self.sheet_tri = self.sheet.loc[self.sheet['Directeur Régional'] == "Cyril ROBINET"]
                print (self.sheet_tri)
                self.model = pandasModel(self.sheet_tri)
            elif self.role == "Magasin":
                ##### Gathering only the lines for a site #####
                self.sheet_tri = self.sheet.loc[self.sheet['Directeur dépôt'] == "YANNICK PIERRE"]
                self.model = pandasModel(self.sheet_tri)
        else:
            self.model = pandasModel(self.sheet)
            print("Model créé Admin")

    def loadGUI(self):
        if self.role == "Admin":
            print("Creation du GUI ADMIN")
            ##### Creation of the admin interface #####

            self.adminGUI = QWidget()

            ##### Top Banner #####
            self.titre = QLabel("Base de donnée magasin - Brico Dépôt")
            self.titre.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.titre.setFont(QFont('Arial', 18))
            self.titre.setStyleSheet("QLabel {background-color: green; color: white;}")
            #self.logo = QPixmap("C:/Solutec/logo.png")
            self.logo = QLabel("IMAGE")
            self.logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.logo.setStyleSheet("QLabel {background-color: red;}")
            self.infos = QLabel("Bandeau d'info \n sur plusieur lignes \n admin")
            self.infos.setAlignment(Qt.AlignmentFlag.AlignRight)
            self.infos.setStyleSheet("QLabel {background-color: blue; margin-right: 50px}")

            self.topBanner = QWidget()
            self.topLayout = QHBoxLayout()
            self.topLayout.addWidget(self.logo)
            self.topLayout.addWidget(self.titre)
            self.topLayout.addWidget(self.infos)

            self.topBanner.setLayout(self.topLayout)

            ##### Table creation #####
            self.middle = QWidget()
            self.middleLayout = QVBoxLayout()
            self.tab = QTableView()
            self.tab.setModel(self.model)
            self.middleLayout.addWidget(self.tab)
            self.middle.setLayout(self.middleLayout)

            ##### Push buttons #####
            self.creer = QPushButton("Créer un dépôt")
            self.creer.clicked.connect(self.creerDepot)
            self.modifier = QPushButton("Modifier un dépôt")
            self.modifier.clicked.connect(self.modifDepot)
            self.supprimer = QPushButton("Supprimer un dépôt")
            self.supprimer.clicked.connect(self.supprimerDepot)

            self.bandeau = QWidget()
            self.bandeauBoutons = QHBoxLayout()
            self.bandeauBoutons.addWidget(self.creer)
            self.bandeauBoutons.addWidget(self.modifier)
            self.bandeauBoutons.addWidget(self.supprimer)

            self.bandeau.setLayout(self.bandeauBoutons)

            ##### Setting up the Widget #####
            self.layoutAdminGUI = QVBoxLayout()
            self.layoutAdminGUI.addWidget(self.topBanner)
            self.layoutAdminGUI.addWidget(self.middle)
            self.layoutAdminGUI.addWidget(self.bandeau)
            self.adminGUI.setLayout(self.layoutAdminGUI)

            ##### Adding the interface to the StackedWidget #####
            self.Stack.addWidget(self.adminGUI)
            self.Stack.setCurrentIndex(1)

        else:
            print("Creation du GUI en lecture")
            ##### Creation of the regular interface #####
            self.adminGUI = QWidget()

            ##### Top Banner #####
            self.titre = QLabel("Base de donnée magasin - Brico Dépôt")
            #self.logo = QPixmap("C:/Solutec/logo.png")
            self.logo = QLabel("IMAGE")
            self.infos = QLabel("Bandeau d'info \n sur plusieur lignes \n role")

            self.topBanner = QWidget()
            self.topLayout = QHBoxLayout()
            self.topLayout.addWidget(self.logo)
            self.topLayout.addWidget(self.titre)
            self.topLayout.addWidget(self.infos)

            self.topBanner.setLayout(self.topLayout)

            ##### Table creation #####
            self.middle = QWidget()
            self.middleLayout = QVBoxLayout()
            self.tab = QTableView()
            self.tab.setModel(self.model)
            self.middleLayout.addWidget(self.tab)
            
            self.middle.setLayout(self.middleLayout)

            ##### Push buttons #####
            self.requete = QPushButton("Demander un changement")

            self.bandeau = QWidget()
            self.bandeauBoutons = QHBoxLayout()
            self.bandeauBoutons.addWidget(self.requete)
            self.bandeau.setLayout(self.bandeauBoutons)

            ##### Setting up the Widget #####
            self.layoutAdminGUI = QVBoxLayout()
            self.layoutAdminGUI.addWidget(self.topBanner)
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
        
    def testModif(self):
        print(self.sheet['Dépôt'][0])
        self.sheet['Dépôt'][0] = "Prout"
        self.tab.setModel(self.model)
        print("Nouvelle valeur:")
        print(self.sheet['Dépôt'][0])

    def creerDepot(self):
        print("Création")

    def modifDepot(self):
        print("Modification")

    def supprimerDepot(self):
        print("Suppression")

class creaWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Création d'un dépôt")
        self.title = QLabel("Création d'un dépôt")
        self.title.setFont(QFont('Arial', 18))
        self.layout = QGridLayout()

        self.ligne1 = QHBoxLayout()

        self.data1 = QLineEdit()
        self.data2 = QLineEdit()
        self.data3 = QLineEdit()
        self.ligne1.addWidget(self.data1)
        self.ligne1.addWidget(self.data2)
        self.ligne1.addWidget(self.data3)

        self.ligne2 = QHBoxLayout()

        self.data4 = QLineEdit()
        self.data5 = QLineEdit()
        self.data6 = QLineEdit()
        self.ligne2.addWidget(self.data4)
        self.ligne2.addWidget(self.data5)
        self.ligne2.addWidget(self.data6)

        self.ligne3 = QHBoxLayout()

        self.data7 = QLineEdit()
        self.data8 = QLineEdit()
        self.data9 = QLineEdit()
        self.ligne3.addWidget(self.data7)
        self.ligne3.addWidget(self.data8)
        self.ligne3.addWidget(self.data9)

        self.boutonCreer = QPushButton("Créer le dépôt")
        self.boutonCreer.clicked.connect(self.insertionbdd)

        self.layout.addWidget(self.title, 0, 0, 1, 3, Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.ligne1, 1, 0, 1, 3)
        self.layout.addWidget(self.ligne2, 2, 0, 1, 3)
        self.layout.addWidget(self.ligne3, 3, 0, 1, 3)
        self.layout.addWidget(self.boutonCreer, 4, 1, 1, 1, Qt.AlignmentFlag.AlignCenter)

        self.setLayout(self.layout)

    def insertionbdd(self):
        print("Bibimbap")

class modifWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Modification d'un dépôt")
        self.title = QLabel("Modification d'un dépôt")

        ##TODO: Faire un stacked widget
        self.stack = QStackedWidget()
                
        ##TODO: Widget 1 -> Sélection du dépot
        self.layout1 = QGridLayout()
        self.listedepots = QComboBox()
        
        ##TODO: Widget 2 -> Affichage du dépot + possibilité de toujours le modifier
        self.layout2 = QGridLayout()
        self.listedepots2 = QComboBox()
        self.affichageDep = QTableView()

        self.boutonConfirmation = QPushButton()
        self.boutonConfirmation.clicked.connect(self.choix)

        ##TODO: Widget 2 -> Accès aux infos du Widget, modification des infos et bouton confirmation
        self.layout3 = QGridLayout()

        self.boutonCreer = QPushButton("Modifier le dépôt")

    def combo(self):
        print("Combo")
        #Charger les données dans self.affichageDep
        self.stack.setCurrentIndex(1)

    def choix(self):
        print("Choxi")
        self.stack.setCurrentIndex(2)

    def modif(self):
        print("Modification faite")
        self.close()


class suppWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Suppression d'un dépôt")

        self.stack = QStackedWidget()
                
        ##TODO: Widget 1 -> Sélection du dépot
        self.layout1 = QGridLayout()
        self.listedepots = QComboBox()
        
        ##TODO: Widget 2 -> Affichage du dépot + possibilité de toujours le modifier
        self.layout2 = QGridLayout()
        self.listedepots2 = QComboBox()
        self.affichageDep = QTableView()

        self.boutonConfirmation = QPushButton()
        self.boutonConfirmation.clicked.connect(self.choix)

        ##TODO: Faire une fenetre de confirmation

    def choix(self):
        print("supp")

class demandeChangement(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Requête changement donnée")

class pandasModel(QAbstractTableModel):
    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if index.isValid():
            if role == Qt.ItemDataRole.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._data.columns[col]
        return None

if __name__ == '__main__':
    app = QApplication(sys.argv)

    form = logWindow()
    form.show()

    sys.exit(app.exec())
