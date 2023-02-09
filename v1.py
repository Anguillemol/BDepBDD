import sys
from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from PyQt6.uic import loadUi

import pandas as pd
import numpy as np
import openpyxl

##TODO: Améliorer passation des données
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
        self.w = creaWindow()
        self.w.show()

    def modifDepot(self):
        print("Modification")
        modifWindow.sheet=self.sheet
        self.w = modifWindow()
        self.w.show()

    def supprimerDepot(self):
        print("Suppression")
        self.w = suppWindow()
        self.w.show()

##TODO: Faire l'insertion des données dans le DataFrame
class creaWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Création d'un dépôt")
        self.title = QLabel("Création d'un dépôt")
        self.title.setFont(QFont('Arial', 18))
        self.layout = QGridLayout()
        
        #Premiere ligne de boutons
        self.r1 = QWidget()
        self.row1 = QHBoxLayout()

        self.listeDepot = QPushButton("Liste dépôt")
        self.listeDepot.clicked.connect(self.ldepot)
        self.w1 = listeDepot()
        self.donneesSociales = QPushButton("Données sociales")
        self.donneesSociales.clicked.connect(self.dSociales)
        self.w2 = donneesSociales()
        self.secteurSecurite = QPushButton("Secteur sécurité")
        self.secteurSecurite.clicked.connect(self.sSecurite)
        self.w3 = secteurSecurite()

        self.row1.addWidget(self.listeDepot)
        self.row1.addWidget(self.donneesSociales)
        self.row1.addWidget(self.secteurSecurite)

        self.r1.setLayout(self.row1)

        #Seconde ligne de boutons
        self.r2 = QWidget()
        self.row2 = QHBoxLayout()

        self.secteurLogistique = QPushButton("Secteur logistique")
        self.secteurLogistique.clicked.connect(self.sLogistique)
        self.w4 = secteurLogistique()
        self.secteurAmenagement = QPushButton("Sectuer aménagement")
        self.secteurAmenagement.clicked.connect(self.sAmenagement)
        self.w5 = secteurAmenagement()
        self.secteurConstruction = QPushButton("Secteur construction")
        self.secteurConstruction.clicked.connect(self.sConstruction)
        self.w6 = secteurConstruction()

        self.row2.addWidget(self.secteurLogistique)
        self.row2.addWidget(self.secteurAmenagement)
        self.row2.addWidget(self.secteurConstruction)
        
        self.r2.setLayout(self.row2)

        #Troisième ligne de boutons
        self.r3 = QWidget()
        self.row3= QHBoxLayout()

        self.secteurTechnique = QPushButton("Secteur technique")
        self.secteurTechnique.clicked.connect(self.sTechnique)
        self.w7 = secteurTechnique()
        self.secteurAdministratif = QPushButton("Secteur administratif")
        self.secteurAdministratif.clicked.connect(self.sAdministratif)
        self.w8 = secteurAdministratif()
        self.secteurCaisse = QPushButton("Secteur caisse")
        self.secteurCaisse.clicked.connect(self.sCaisse)
        self.w9 = secteurCaisse()

        self.row3.addWidget(self.secteurTechnique)
        self.row3.addWidget(self.secteurAdministratif)
        self.row3.addWidget(self.secteurCaisse)

        self.r3.setLayout(self.row3)

        #Quatrième ligne de boutons
        self.r4 = QWidget()
        self.row4 = QHBoxLayout()

        self.RH = QPushButton("RH")
        self.RH.clicked.connect(self.RHFonction)
        self.w10 = RH()
        self.donneesDepot = QPushButton("Données dépôt")
        self.donneesDepot.clicked.connect(self.dDepot)
        self.w11 = donneesDepot()
        self.surface = QPushButton("Surface")
        self.surface.clicked.connect(self.surfaceFonction)
        self.w12 = surface()

        self.row4.addWidget(self.RH)
        self.row4.addWidget(self.donneesDepot)
        self.row4.addWidget(self.surface)

        self.r4.setLayout(self.row4)

        #Cinquième ligne de boutons
        self.r5 = QWidget()
        self.row5 = QHBoxLayout()

        self.agencement = QPushButton("Agencement")
        self.agencement.clicked.connect(self.agencementFonction)
        self.w13 = agencement()
        self.caisse = QPushButton("Caisse")
        self.caisse.clicked.connect(self.caisseFonction)
        self.w14 = caisse()
        self.PDA = QPushButton("PDA")
        self.PDA.clicked.connect(self.PDAFonction)
        self.w15 = PDA()

        self.row5.addWidget(self.agencement)
        self.row5.addWidget(self.caisse)
        self.row5.addWidget(self.PDA)

        self.r5.setLayout(self.row5)

        #Sixième ligne de boutons
        self.r6 = QWidget()
        self.row6 = QHBoxLayout()

        self.menace = QPushButton("Menace")
        self.menace.clicked.connect(self.menaceFonction)
        self.w16 = menace()
        self.securite = QPushButton("Sécurité")
        self.securite.clicked.connect(self.securiteFonction)
        self.w17 = securite()
        self.conceptCommercial = QPushButton("Concept commercial")
        self.conceptCommercial.clicked.connect(self.cCommercial)
        self.w18 = conceptCommercial()

        self.row6.addWidget(self.menace)
        self.row6.addWidget(self.securite)
        self.row6.addWidget(self.conceptCommercial)

        self.r6.setLayout(self.row6)

        #Septième ligne de boutons
        self.r7 = QWidget()
        self.row7 = QHBoxLayout()

        self.divers = QPushButton("Divers")
        self.divers.clicked.connect(self.diversFonction)
        self.w19 = divers()
        self.numCommercant = QPushButton("Num commercant")
        self.numCommercant.clicked.connect(self.nCommercant)
        self.w20 = numCommercant()
        self.colissimo = QPushButton("Colissimo")
        self.colissimo.clicked.connect(self.colissimoFonction)
        self.w21 = colissimo()

        self.row7.addWidget(self.divers)
        self.row7.addWidget(self.numCommercant)
        self.row7.addWidget(self.colissimo)

        self.r7.setLayout(self.row7)

        #Huitième ligne de bouton
        self.r8 = QWidget()
        self.row8 = QHBoxLayout()

        self.accidentTravail = QPushButton("Accident travail")
        self.accidentTravail.clicked.connect(self.aTravail)
        self.w22 = accidentTravail()

        self.row8.addWidget(self.accidentTravail)

        self.r8.setLayout(self.row8)

        self.boutonConfirmer = QPushButton("Confirmer la création du dépôt")
        self.boutonConfirmer.clicked.connect(self.confirmerCreation)

        self.layout.addWidget(self.title, 0, 0, Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.r1)
        self.layout.addWidget(self.r2)
        self.layout.addWidget(self.r3)
        self.layout.addWidget(self.r4)
        self.layout.addWidget(self.r5)
        self.layout.addWidget(self.r6)
        self.layout.addWidget(self.r7)
        self.layout.addWidget(self.r8)
        self.layout.addWidget(self.boutonConfirmer)

        self.setLayout(self.layout)

    def ldepot(self):
        print("Liste dépôt")
        self.w = self.w1
        self.w.show()

    def dSociales(self):
        print("Données sociales")
        self.w = self.w2
        self.w.show()

    def sSecurite(self):
        print("Secteur sécurité")
        self.w = self.w3
        self.w.show()

    def sLogistique(self):
        print("")

    def sAmenagement(self):
        print("")

    def sConstruction(self):
        print("")

    def sTechnique(self):
        print("")
    
    def sAdministratif(self):
        print("")

    def sCaisse(self):
        print("")

    def RHFonction(self):
        print("")

    def dDepot(self):
        print("")

    def surfaceFonction(self):
        print("")

    def agencementFonction(self):
        print("")

    def caisseFonction(self):
        print("")

    def PDAFonction(self):
        print("")

    def menaceFonction(self):
        print("")

    def securiteFonction(self):
        print("")

    def cCommercial(self):
        print("")

    def diversFonction(self):
        print("")

    def nCommercant(self):
        print("")    

    def colissimoFonction(self):
        print("")   

    def aTravail(self):
        print("")   

    def confirmerCreation(self):
        print("Confirmation")
        print(self.w1.test.text())
        #Cehcker si tt les données sont entrées


class listeDepot(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle ("Liste dépôt")
        self.layout = QHBoxLayout()
        self.titre = QLabel("Liste Dépot")
        self.test = QLineEdit()

        self.layout.addWidget(self.titre)
        self.layout.addWidget(self.test)
        self.setLayout(self.layout)

class donneesSociales(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Données sociales")

class secteurSecurite(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur sécurité")

class secteurLogistique(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur logistique")

class secteurAmenagement(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur amenagement")

class secteurConstruction(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur construction")

class secteurTechnique(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur technique")

class secteurAdministratif(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur administratif")

class secteurCaisse(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur caisse")

class RH(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RH")

class donneesDepot(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Données dépôt")

class surface(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Surface"
)
class agencement(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Agencement")

class caisse(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Caisse")

class PDA(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDA")

class menace(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Menace")

class securite(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sécurité")

class conceptCommercial(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Concept commercial")

class divers(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Divers")

class numCommercant(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Numéro commercant")

class colissimo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Colissimo")

class accidentTravail(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Accident travail")

##TODO: Gérer le dimensionnement des fenêtres ou le placement statique
class modifWindow(QWidget):
    sheet = pd.DataFrame

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Modification d'un dépôt")
        self.title = QLabel("Modification d'un dépôt")
        self.setFixedSize(200,80)

        ##TODO: Faire un stacked widget
        self.stack = QStackedWidget()
                
        ##TODO: Widget 1 -> Sélection du dépot
        self.widget1 = QWidget()
        self.layout1 = QVBoxLayout()

        self.titre = QLabel("Sélection du dépôt à modifier")
        self.listedepots = QComboBox()
        self.listedesdepots = self.sheet["Dépôt"].values.tolist()
        self.listedepots.addItems(self.listedesdepots)
        self.listedepots.currentIndexChanged.connect(self.selectionDepot)



        self.layout1.addWidget(self.titre)
        self.layout1.addWidget(self.listedepots)

        self.widget1.setLayout(self.layout1)
        
        ##TODO: Widget 2 -> Affichage du dépot + possibilité de toujours le modifier
        self.widget2 = QWidget()
        self.layout2 = QGridLayout()
        self.titre2 = QLabel("Sélection du dépôt à modifier")
        self.listedepots2 = QComboBox()
        self.listedepots2.addItems(self.listedesdepots)
        self.affichageDep = QTableView()

        self.boutonConfirmation = QPushButton()
        self.boutonConfirmation.clicked.connect(self.choix)
        #, Qt.AlignmentFlag.AlignCenter
        self.layout2.addWidget(self.titre2, 0, 1)
        self.layout2.addWidget(self.listedepots2, 1, 1,)
        self.layout2.addWidget(self.affichageDep, 2, 0, 1, 3)
        self.layout2.addWidget(self.boutonConfirmation, 3, 1)


        self.widget2.setLayout(self.layout2)

        ##TODO: Widget 2 -> Accès aux infos du Widget, modification des infos et bouton confirmation
        self.widget3 = QWidget()
        self.layout3 = QGridLayout()

        self.boutonCreer = QPushButton("Modifier le dépôt")
        self.boutonCreer.clicked.connect(self.modif)
        self.layout3.addWidget(self.boutonCreer)

        self.widget3.setLayout(self.layout3)

        self.mainlayout = QGridLayout()
        self.stack.addWidget(self.widget1)
        self.stack.addWidget(self.widget2)
        self.stack.addWidget(self.widget3)
        self.mainlayout.addWidget(self.stack)
        self.setLayout(self.mainlayout)

    def selectionDepot(self):
        ##Transition vers la seconde fenêtre
        self.stack.setCurrentIndex(1)    
        self.setFixedSize(720,440)
    
    def choix(self):
        ##Transition vers la fenêtre de choix des modifications
        self.stack.setCurrentIndex(2)

    def modif(self):
        ##Application des modifications
        print("Modification faite")
        ##Fermeture de la fenetre et refresh de la dataframe
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
