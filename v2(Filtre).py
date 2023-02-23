import sys
from pathlib import Path
from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from PyQt6.uic import loadUi

import io

from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File 

import pandas as pd
import numpy as np

username = '50bbb53b-67ef-488d-9303-d6afcfd77bc8'
password = '7ATT8OvZyqU1jbWSFxsgiDMZXrqJ4KekP/JMkgFRQCc='

test_team_site_url = "https://sgzkl.sharepoint.com/sites/SiteTest"


ctx = ClientContext(test_team_site_url).with_credentials(ClientCredential(username, password))
bdd_URL = "/sites/SiteTest/Documents%20partages/Test/BDD.xlsx"

mdp_URL = "/sites/SiteTest/Documents%20partages/Test/MDP.xlsx"

requete_URL = "/sites/SiteTest/Documents%20partages/Test/REQ.xlsx"

response = File.open_binary(ctx, bdd_URL)

bytes_file_obj_bdd = io.BytesIO()
bytes_file_obj_bdd.write(response.content)
bytes_file_obj_bdd.seek(0)

response = File.open_binary(ctx, mdp_URL)

bytes_file_obj_mdp = io.BytesIO()
bytes_file_obj_mdp.write(response.content)
bytes_file_obj_mdp.seek(0)

response = File.open_binary(ctx, requete_URL)

bytes_file_obj_req = io.BytesIO()
bytes_file_obj_req.write(response.content)
bytes_file_obj_req.seek(0)

#Read
bdd = bytes_file_obj_bdd
mdp = bytes_file_obj_mdp
req = bytes_file_obj_req




p = str(Path.cwd())
p = p.replace('\\', "/")

##TODO: Faire l'ouverture du fichier Excel sur Sharepoint
class logWindow(QWidget):

    
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Authentification BricoDepot")
        self.w = None 
        self.role = ""
        self.submitClicked = pyqtSignal(str,str,str)

        ########## Gathering the account DataBase ##########
     
        #self.ddbmdp = pd.read_excel(p+'/Test.xlsx', sheet_name='MDPFINAUX')  

        self.ddbmdp = pd.read_excel(mdp, sheet_name='MDPFINAUX')
        
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

##TODO: Rajouter un bouton de validation et la fonction qui enregistre le Dataframe sur Sharepoint en répartissant dans tous les onglets (rajouter nom de l'onglet dans chaque classe)
class mainWindow(QWidget):


    def __init__(self):
        self.submitClicked = pyqtSignal(str,str,str)
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
        #TODO: Recentrer la fenêtre au milieu de l'écran
        
        ##### Generation of the dataFrame #####
        self.loadExcel()

        ##### Generation of the layout #####
        self.loadGUI()
        
    ##### Function used to load the Excel data file #####
    def loadExcel(self):
        self.sheet = pd.read_excel(p+'/BDD.xlsx', sheet_name='BDD')

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

            self.searchBar = QLineEdit()

            self.tab = QTableView()

            self.proxy_model = QSortFilterProxyModel()
            self.proxy_model.setFilterKeyColumn(-1) #Toutes les colonnes
            self.proxy_model.setSourceModel(self.model)

            self.tab.setModel(self.proxy_model)
            #self.tab.setModel(self.model)
            self.tab.resizeColumnsToContents()

            self.searchBar.textChanged.connect(self.proxy_model.setFilterFixedString)

            self.middleLayout.addWidget(self.searchBar)
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
            self.tab.resizeColumnsToContents()
            self.middleLayout.addWidget(self.tab)
            
            self.middle.setLayout(self.middleLayout)

            ##### Push buttons #####
            self.requete = QPushButton("Demander un changement")
            self.requete.clicked.connect(self.demandeChangement)

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
        self.tab.resizeColumnsToContents()
        print("Nouvelle valeur:")
        print(self.sheet['Dépôt'][0])

    def creerDepot(self):
        print("Création")
        creaWindow.sheet = self.sheet
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

    def chargerModif(self):
        self.model = pandasModel(self.sheet)
        self.tab.setModel(self.model)
        self.tab.resizeColumnsToContents()
        print("Model chargé")


    def closeEvent(self, event):
        self.w = ''
        if self.w:
            self.w.close()

##TODO: Faire l'insertion des données dans le DataFrame concat()
##TODO: Prendre toutes les données de toutes les classes et les poser dans un df d'1 ligne
class creaWindow(QWidget):
    sheet = pd.DataFrame
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
        self.w = self.w4
        self.w.show()

    def sAmenagement(self):
        self.w = self.w5
        self.w.show()

    def sConstruction(self):
        self.w = self.w6
        self.w.show()

    def sTechnique(self):
        self.w = self.w7
        self.w.show()

    def sAdministratif(self):
        self.w = self.w8
        self.w.show()

    def sCaisse(self):
        self.w = self.w9
        self.w.show()

    def RHFonction(self):
        self.w = self.w10
        self.w.show()

    def dDepot(self):
        self.w = self.w11
        self.w.show()

    def surfaceFonction(self):
        self.w = self.w12
        self.w.show()

    def agencementFonction(self):
        self.w = self.w13
        self.w.show()

    def caisseFonction(self):
        self.w = self.w14
        self.w.show()

    def PDAFonction(self):
        self.w = self.w15
        self.w.show()

    def menaceFonction(self):
        self.w = self.w16
        self.w.show()

    def securiteFonction(self):
        self.w = self.w17
        self.w.show()

    def cCommercial(self):
        self.w = self.w18
        self.w.show()

    def diversFonction(self):
        self.w = self.w19
        self.w.show()

    def nCommercant(self):
        self.w = self.w20
        self.w.show()    

    def colissimoFonction(self):
        self.w = self.w21
        self.w.show()   

    def aTravail(self):
        self.w = self.w22
        self.w.show()   

    def confirmerCreation(self):
        print("Confirmation")

        
        my_dict = {}
        my_dict[self.w1.labelcodeBRICO.text()] = self.w1.codeBRICO.text()
        my_dict[self.w1.labelcodeEASIER.text()] = self.w1.codeEASIER.text()
        my_dict[self.w1.labeldepot.text()] = self.w1.depot.text()
        my_dict[self.w1.labelregionAdministrative.text()] = self.w1.regionAdministrative.text()
        my_dict[self.w1.labelregion2022.text()] = self.w1.region2022.text()
        my_dict[self.w1.labelstatut.text()] = self.w1.statut.text()
        my_dict[self.w2.labelDirecteurRegional.text()] = self.w2.DirecteurRegional.text()
        my_dict[self.w2.labelDirecteurRegionaltelephone.text()] = self.w2.DirecteurRegionaltelephone.text()
        my_dict[self.w2.labelDirecteurdepot.text()] = self.w2.Directeurdepot.text()
        my_dict[self.w2.labelDirecteurdepotmail.text()] = self.w2.Directeurdepotmail.text()
        my_dict[self.w2.labelDirecteurdepottelephone.text()] = self.w2.Directeurdepottelephone.text()
        my_dict[self.w2.labelDatedeprisedefonction.text()] = self.w2.Datedeprisedefonction.text()
        my_dict[self.w3.labelChefSecurite.text()] = self.w3.ChefSecurite.text()
        my_dict[self.w3.labelChefSecuritemail.text()] = self.w3.ChefSecuritemail.text()
        my_dict[self.w3.labelChefSecuriteTelephone.text()] = self.w3.ChefSecuriteTelephone.text()
        my_dict[self.w3.labelChefSecuriteStatut.text()] = self.w3.ChefSecuriteStatut.text()
        my_dict[self.w3.labelDMfileSecurite.text()] = self.w3.DMfileSecurite.text()
        my_dict[self.w3.labelCSReferentSecurite.text()] = self.w3.CSReferentSecurite.text()
        my_dict[self.w4.labelChefLogistique.text()] = self.w4.ChefLogistique.text()
        my_dict[self.w4.labelChefLogistiquemail.text()] = self.w4.ChefLogistiquemail.text()
        my_dict[self.w4.labelChefLogistiqueTelephone.text()] = self.w4.ChefLogistiqueTelephone.text()
        my_dict[self.w4.labelChefLogistiqueStatut.text()] = self.w4.ChefLogistiqueStatut.text()
        my_dict[self.w4.labelDMfileLogistique.text()] = self.w4.DMfileLogistique.text()
        my_dict[self.w4.labelCSReferentLogistique.text()] = self.w4.CSReferentLogistique.text()
        my_dict[self.w5.labelChefAmenagement.text()] = self.w5.ChefAmenagement.text()
        my_dict[self.w5.labelChefAmenagementmail.text()] = self.w5.ChefAmenagementmail.text()
        my_dict[self.w5.labelChefAmenagementTelephone.text()] = self.w5.ChefAmenagementTelephone.text()
        my_dict[self.w5.labelDMfileAmenagement.text()] = self.w5.DMfileAmenagement.text()
        my_dict[self.w5.labelCSReferentAmenagement.text()] = self.w5.CSReferentAmenagement.text()
        my_dict[self.w6.labelChefConstruction.text()] = self.w6.ChefConstruction.text()
        my_dict[self.w6.labelChefConstructionmail.text()] = self.w6.ChefConstructionmail.text()
        my_dict[self.w6.labelChefConstructionTelephone.text()] = self.w6.ChefConstructionTelephone.text()
        my_dict[self.w6.labelDMfileConstruction.text()] = self.w6.DMfileConstruction.text()
        my_dict[self.w6.labelCSReferentConstruction.text()] = self.w6.CSReferentConstruction.text()
        my_dict[self.w7.labelChefTechnique.text()] = self.w7.ChefTechnique.text()
        my_dict[self.w7.labelChefTechniquemail.text()] = self.w7.ChefTechniquemail.text()
        my_dict[self.w7.labelChefTechniqueTelephone.text()] = self.w7.ChefTechniqueTelephone.text()
        my_dict[self.w7.labelDMfileTechnique.text()] = self.w7.DMfileTechnique.text()
        my_dict[self.w7.labelCSReferentTechnique.text()] = self.w7.CSReferentTechnique.text()
        my_dict[self.w8.labelChefAdministratif.text()] = self.w8.ChefAdministratif.text()
        my_dict[self.w8.labelChefAdministratifmail.text()] = self.w8.ChefAdministratifmail.text()
        my_dict[self.w8.labelChefAdministratifTelephone.text()] = self.w8.ChefAdministratifTelephone.text()
        my_dict[self.w8.labelDMfileAdministratif.text()] = self.w8.DMfileAdministratif.text()
        my_dict[self.w8.labelCSReferentAdministratif.text()] = self.w8.CSReferentAdministratif.text()
        my_dict[self.w9.labelChefCaisse.text()] = self.w9.ChefCaisse.text()
        my_dict[self.w9.labelChefCaissemail.text()] = self.w9.ChefCaissemail.text()
        my_dict[self.w9.labelChefCaisseTelephone.text()] = self.w9.ChefCaisseTelephone.text()
        my_dict[self.w9.labelDMfileCaisse.text()] = self.w9.DMfileCaisse.text()
        my_dict[self.w9.labelCSReferentCaisse.text()] = self.w9.CSReferentCaisse.text()
        my_dict[self.w10.labelResponsableRH.text()] = self.w10.ResponsableRH.text()
        my_dict[self.w10.labelResponsableRHTelephone.text()] = self.w10.ResponsableRHTelephone.text()
        my_dict[self.w10.labelResponsableCDG.text()] = self.w10.ResponsableCDG.text()
        my_dict[self.w10.labelResponsableCDGTelephone.text()] = self.w10.ResponsableCDGTelephone.text()
        my_dict[self.w11.labelAdresse.text()] = self.w11.Adresse.text()
        my_dict[self.w11.labelVille.text()] = self.w11.Ville.text()
        my_dict[self.w11.labelDepartement.text()] = self.w11.Departement.text()
        my_dict[self.w11.labelTELEPHONESTANDARD.text()] = self.w11.TELEPHONESTANDARD.text()
        my_dict[self.w11.labelLatitudeendegresdecimaux.text()] = self.w11.Latitudeendegresdecimaux.text()
        my_dict[self.w11.labelLongitudeendegresdecimaux.text()] = self.w11.Longitudeendegresdecimaux.text()
        my_dict[self.w11.labelNdegSIREN.text()] = self.w11.NdegSIREN.text()
        my_dict[self.w11.labelSIRET.text()] = self.w11.SIRET.text()
        my_dict[self.w11.labelDatedOuverture.text()] = self.w11.DatedOuverture.text()
        my_dict[self.w11.labelAnciennetedudepot.text()] = self.w11.Anciennetedudepot.text()
        my_dict[self.w11.labelHorairedouverturejours1.text()] = self.w11.Horairedouverturejours1.text()
        my_dict[self.w11.labelHorairedouverturehoraires1.text()] = self.w11.Horairedouverturehoraires1.text()
        my_dict[self.w11.labelHorairedouverturejours2.text()] = self.w11.Horairedouverturejours2.text()
        my_dict[self.w11.labelHorairedouverturehoraires2.text()] = self.w11.Horairedouverturehoraires2.text()
        my_dict[self.w11.labelHorairedouverturejours3.text()] = self.w11.Horairedouverturejours3.text()
        my_dict[self.w11.labelHorairedouverturehoraires3.text()] = self.w11.Horairedouverturehoraires3.text()
        my_dict[self.w11.labelAmplitudeHoraire.text()] = self.w11.AmplitudeHoraire.text()
        my_dict[self.w11.labelCATTC2021.text()] = self.w11.CATTC2021.text()
        my_dict[self.w11.labelBASSIN2021.text()] = self.w11.BASSIN2021.text()
        my_dict[self.w11.labelCluster2021.text()] = self.w11.Cluster2021.text()
        my_dict[self.w11.labelPassagesCaisses2021.text()] = self.w11.PassagesCaisses2021.text()
        my_dict[self.w11.labelTauxdeDemarque2021.text()] = self.w11.TauxdeDemarque2021.text()
        my_dict[self.w11.labelETPCDI2021.text()] = self.w11.ETPCDI2021.text()
        my_dict[self.w11.labelETPCDD2021.text()] = self.w11.ETPCDD2021.text()
        my_dict[self.w11.labelETPINTERIM2021.text()] = self.w11.ETPINTERIM2021.text()
        my_dict[self.w11.labelETPGLOBAL2021.text()] = self.w11.ETPGLOBAL2021.text()
        my_dict[self.w11.labelFormatdepot.text()] = self.w11.Formatdepot.text()
        my_dict[self.w11.labelNombredesites.text()] = self.w11.Nombredesites.text()
        my_dict[self.w11.labelZoneScolaire.text()] = self.w11.ZoneScolaire.text()
        my_dict[self.w11.labelZonedeDroitSocial.text()] = self.w11.ZonedeDroitSocial.text()
        my_dict[self.w11.labelEntrepotLegacy.text()] = self.w11.EntrepotLegacy.text()
        my_dict[self.w11.labelEntrepotEasier.text()] = self.w11.EntrepotEasier.text()
        my_dict[self.w11.labelPlateformeLegacy.text()] = self.w11.PlateformeLegacy.text()
        my_dict[self.w11.labelPlateformeEasier.text()] = self.w11.PlateformeEasier.text()
        my_dict[self.w11.labelPlateformeArrivage.text()] = self.w11.PlateformeArrivage.text()
        my_dict[self.w11.labelPlateformeArrivageECC.text()] = self.w11.PlateformeArrivageECC.text()
        my_dict[self.w11.labelPlateformederattachementGammeetArrivagesapresouvertures.text()] = self.w11.PlateformederattachementGammeetArrivagesapresouvertures.text()
        my_dict[self.w11.labelGLNPlateformeGamme.text()] = self.w11.GLNPlateformeGamme.text()
        my_dict[self.w11.labelCodeplateforme.text()] = self.w11.Codeplateforme.text()
        my_dict[self.w11.labelEntrepotdeporte.text()] = self.w11.Entrepotdeporte.text()
        my_dict[self.w11.labelOrigineduDepot.text()] = self.w11.OrigineduDepot.text()
        my_dict[self.w11.labelSituationduDepot.text()] = self.w11.SituationduDepot.text()
        my_dict[self.w11.labelProprietaireouLocataire.text()] = self.w11.ProprietaireouLocataire.text()
        my_dict[self.w12.labelSurfaceTotaleCDAC.text()] = self.w12.SurfaceTotaleCDAC.text()
        my_dict[self.w12.labelSurfacedeventeinterieure.text()] = self.w12.Surfacedeventeinterieure.text()
        my_dict[self.w12.labelTypedesurfaceSVI.text()] = self.w12.TypedesurfaceSVI.text()
        my_dict[self.w12.labelTypologieSVI.text()] = self.w12.TypologieSVI.text()
        my_dict[self.w12.labelCourMateriaux.text()] = self.w12.CourMateriaux.text()
        my_dict[self.w12.labelDistanceCourMateriauxDepot.text()] = self.w12.DistanceCourMateriauxDepot.text()
        my_dict[self.w12.labelSurfaceBati.text()] = self.w12.SurfaceBati.text()
        my_dict[self.w12.labelTypedesurfaceBati.text()] = self.w12.TypedesurfaceBati.text()
        my_dict[self.w12.labelTypologieBati.text()] = self.w12.TypologieBati.text()
        my_dict[self.w12.labelSurfacedesmateriauxenCDAC.text()] = self.w12.SurfacedesmateriauxenCDAC.text()
        my_dict[self.w12.labelSURFACEBatienCDAC.text()] = self.w12.SURFACEBatienCDAC.text()
        my_dict[self.w12.labeldontLocaldeventeBatienCDAC.text()] = self.w12.dontLocaldeventeBatienCDAC.text()
        my_dict[self.w12.labeldontGNBensurfacedevente.text()] = self.w12.dontGNBensurfacedevente.text()
        my_dict[self.w12.labeldontBatiCouvertenCDAC.text()] = self.w12.dontBatiCouvertenCDAC.text()
        my_dict[self.w12.labeldontBatiNoncouvertenCDAC.text()] = self.w12.dontBatiNoncouvertenCDAC.text()
        my_dict[self.w12.labelSurfacedelacourmateriauxhorsCDAC.text()] = self.w12.SurfacedelacourmateriauxhorsCDAC.text()
        my_dict[self.w12.labeldontBatiCouverthorsCDAC.text()] = self.w12.dontBatiCouverthorsCDAC.text()
        my_dict[self.w12.labeldontBatiNonCouverthorsCDAC.text()] = self.w12.dontBatiNonCouverthorsCDAC.text()
        my_dict[self.w12.labelEmplacementMenuiserie.text()] = self.w12.EmplacementMenuiserie.text()
        my_dict[self.w12.labelConfigurationMenuiserie.text()] = self.w12.ConfigurationMenuiserie.text()
        my_dict[self.w12.labelDistanceMenuiserieDepot.text()] = self.w12.DistanceMenuiserieDepot.text()
        my_dict[self.w12.labelShowRoomMenuiserie.text()] = self.w12.ShowRoomMenuiserie.text()
        my_dict[self.w12.labelSurfacedelamenuiserie.text()] = self.w12.Surfacedelamenuiserie.text()
        my_dict[self.w12.labeldontMenuiserieensurfacedevente.text()] = self.w12.dontMenuiserieensurfacedevente.text()
        my_dict[self.w12.labeldontMenuiserieenreserve.text()] = self.w12.dontMenuiserieenreserve.text()
        my_dict[self.w12.labelShowRoomSalledeBains.text()] = self.w12.ShowRoomSalledeBains.text()
        my_dict[self.w12.labelSurfacedelareserve.text()] = self.w12.Surfacedelareserve.text()
        my_dict[self.w12.labelSurfaceduSas.text()] = self.w12.SurfaceduSas.text()
        my_dict[self.w12.labelSurfacedesbureaux.text()] = self.w12.Surfacedesbureaux.text()
        my_dict[self.w12.labelPlacesdeparking.text()] = self.w12.Placesdeparking.text()
        my_dict[self.w13.labelFournisseurAgencement.text()] = self.w13.FournisseurAgencement.text()
        my_dict[self.w13.labelNbredenginsdemanutention.text()] = self.w13.Nbredenginsdemanutention.text()
        my_dict[self.w13.labelSystemedeGestiondeFlotte.text()] = self.w13.SystemedeGestiondeFlotte.text()
        my_dict[self.w14.labelPCDediesFlexpoint.text()] = self.w14.PCDediesFlexpoint.text()
        my_dict[self.w14.labelNbredecaissesACCUEIL.text()] = self.w14.NbredecaissesACCUEIL.text()
        my_dict[self.w14.labelNbredecaissesMAGASIN.text()] = self.w14.NbredecaissesMAGASIN.text()
        my_dict[self.w14.labelNbredeCaissesBATI.text()] = self.w14.NbredeCaissesBATI.text()
        my_dict[self.w14.labelNbredeCaissesGNB.text()] = self.w14.NbredeCaissesGNB.text()
        my_dict[self.w14.labelNbredeCaissesDEPOUILL.text()] = self.w14.NbredeCaissesDEPOUILL.text()
        my_dict[self.w14.labelTotalCaisses.text()] = self.w14.TotalCaisses.text()
        my_dict[self.w14.labelDateRemplacementCaisses.text()] = self.w14.DateRemplacementCaisses.text()
        my_dict[self.w14.labelNbreSCO.text()] = self.w14.NbreSCO.text()
        my_dict[self.w14.labelModeleSCO.text()] = self.w14.ModeleSCO.text()
        my_dict[self.w14.labelModeledeTPE.text()] = self.w14.ModeledeTPE.text()
        my_dict[self.w14.labelDateRemplacementTPE.text()] = self.w14.DateRemplacementTPE.text()
        my_dict[self.w15.labelDotationPDAMotorolaMC75.text()] = self.w15.DotationPDAMotorolaMC75.text()
        my_dict[self.w15.labelPDAMC75Restant012017.text()] = self.w15.PDAMC75Restant012017.text()
        my_dict[self.w15.labelImprimanteZebra.text()] = self.w15.ImprimanteZebra.text()
        my_dict[self.w15.labelPDARelevesdeprix.text()] = self.w15.PDARelevesdeprix.text()
        my_dict[self.w15.labelRescencemntPDA2019.text()] = self.w15.RescencemntPDA2019.text()
        my_dict[self.w16.labelTauxdevolspour100000habitants.text()] = self.w16.Tauxdevolspour100000habitants.text()
        my_dict[self.w16.labelDepotarisquedeVoletnaturedurisque.text()] = self.w16.DepotarisquedeVoletnaturedurisque.text()
        my_dict[self.w16.labelTauxdeBraquagepour100000habitants.text()] = self.w16.TauxdeBraquagepour100000habitants.text()
        my_dict[self.w16.labelBraquages.text()] = self.w16.Braquages.text()
        my_dict[self.w16.labelBraquagecible.text()] = self.w16.Braquagecible.text()
        my_dict[self.w16.labelTauxdecambriolagepour100000habitants.text()] = self.w16.Tauxdecambriolagepour100000habitants.text()
        my_dict[self.w16.labelCambriolages.text()] = self.w16.Cambriolages.text()
        my_dict[self.w16.labelDontCambriolageducoffre.text()] = self.w16.DontCambriolageducoffre.text()
        my_dict[self.w16.labelTotalEvenements.text()] = self.w16.TotalEvenements.text()
        my_dict[self.w17.labelTypedeCoffre.text()] = self.w17.TypedeCoffre.text()
        my_dict[self.w17.labelAutomateGLORY.text()] = self.w17.AutomateGLORY.text()
        my_dict[self.w17.labelModeledeCoffre.text()] = self.w17.ModeledeCoffre.text()
        my_dict[self.w17.labelNombredePointsdeRamassageLOOMIS.text()] = self.w17.NombredePointsdeRamassageLOOMIS.text()
        my_dict[self.w17.labelPneumatiqueCoffre.text()] = self.w17.PneumatiqueCoffre.text()
        my_dict[self.w17.labelAntennesAntiVol.text()] = self.w17.AntennesAntiVol.text()
        my_dict[self.w17.labelVideoCompleteenDepot.text()] = self.w17.VideoCompleteenDepot.text()
        my_dict[self.w17.labelVideodansleslocauxsociaux.text()] = self.w17.Videodansleslocauxsociaux.text()
        my_dict[self.w17.labelVideoenCaisses.text()] = self.w17.VideoenCaisses.text()
        my_dict[self.w17.labelVideoenLogistique.text()] = self.w17.VideoenLogistique.text()
        my_dict[self.w17.labelVideoauBati.text()] = self.w17.VideoauBati.text()
        my_dict[self.w17.labelSocietedeSecurite2019.text()] = self.w17.SocietedeSecurite2019.text()
        my_dict[self.w17.labelCouthoraire2019.text()] = self.w17.Couthoraire2019.text()
        my_dict[self.w17.labelNombredheuresGardiennage2017.text()] = self.w17.NombredheuresGardiennage2017.text()
        my_dict[self.w17.labelSocietedinterventiondeNuit.text()] = self.w17.SocietedinterventiondeNuit.text()
        my_dict[self.w17.labelDepotSprinkle.text()] = self.w17.DepotSprinkle.text()
        my_dict[self.w17.labelSystemedeSecuriteIncendie.text()] = self.w17.SystemedeSecuriteIncendie.text()
        my_dict[self.w17.labelModeledeCentraledAlarme.text()] = self.w17.ModeledeCentraledAlarme.text()
        my_dict[self.w17.labelInstallateurAlarme1.text()] = self.w17.InstallateurAlarme1.text()
        my_dict[self.w17.labelInstallateurAlarme2.text()] = self.w17.InstallateurAlarme2.text()
        my_dict[self.w17.labelTelesurveilleur.text()] = self.w17.Telesurveilleur.text()
        my_dict[self.w17.labelTelesurveilleurTelephone.text()] = self.w17.TelesurveilleurTelephone.text()
        my_dict[self.w17.labelGTB.text()] = self.w17.GTB.text()
        my_dict[self.w17.labelPortedaccesauxlocauxSociaux.text()] = self.w17.PortedaccesauxlocauxSociaux.text()
        my_dict[self.w17.labelControledaccesparbadge.text()] = self.w17.Controledaccesparbadge.text()
        my_dict[self.w17.labelControledaccesMigresurserveur.text()] = self.w17.ControledaccesMigresurserveur.text()
        my_dict[self.w17.labelDATIenMenuiserie.text()] = self.w17.DATIenMenuiserie.text()
        my_dict[self.w17.labelGenerateurdefumeedansleLocalSecurite.text()] = self.w17.GenerateurdefumeedansleLocalSecurite.text()
        my_dict[self.w17.labelRalentisseurssurParking.text()] = self.w17.RalentisseurssurParking.text()
        my_dict[self.w17.labelPlotsantibeliersurParking.text()] = self.w17.PlotsantibeliersurParking.text()
        my_dict[self.w17.labelBarrieresdefermetureduParking.text()] = self.w17.BarrieresdefermetureduParking.text()
        my_dict[self.w17.labelBavoletsLogistique.text()] = self.w17.BavoletsLogistique.text()
        my_dict[self.w17.labelEclairageLED.text()] = self.w17.EclairageLED.text()
        my_dict[self.w17.labelAscenseurs.text()] = self.w17.Ascenseurs.text()
        my_dict[self.w18.labelFacadeROUGE.text()] = self.w18.FacadeROUGE.text()
        my_dict[self.w18.labelRevitalisationRemodling.text()] = self.w18.RevitalisationRemodling.text()
        my_dict[self.w18.labelDRIVE.text()] = self.w18.DRIVE.text()
        my_dict[self.w18.labelDRIVEdate.text()] = self.w18.DRIVEdate.text()
        my_dict[self.w18.labelIDepot.text()] = self.w18.IDepot.text()
        my_dict[self.w18.labelLivraisonLocaleInStore.text()] = self.w18.LivraisonLocaleInStore.text()
        my_dict[self.w18.labelLivraisonLocaleInStoreDate.text()] = self.w18.LivraisonLocaleInStoreDate.text()
        my_dict[self.w18.labelZRM.text()] = self.w18.ZRM.text()
        my_dict[self.w18.labelZRMDate.text()] = self.w18.ZRMDate.text()
        my_dict[self.w18.labelContenuZRM.text()] = self.w18.ContenuZRM.text()
        my_dict[self.w18.labelimprimantededieeZRM.text()] = self.w18.imprimantededieeZRM.text()
        my_dict[self.w18.labelEASIEROULEGACY.text()] = self.w18.EASIEROULEGACY.text()
        my_dict[self.w18.labelBacasable.text()] = self.w18.Bacasable.text()
        my_dict[self.w18.labelFournisseurSable.text()] = self.w18.FournisseurSable.text()
        my_dict[self.w18.labelTremie.text()] = self.w18.Tremie.text()
        my_dict[self.w18.labelGodethydraulique.text()] = self.w18.Godethydraulique.text()
        my_dict[self.w18.labelGodetmecanique.text()] = self.w18.Godetmecanique.text()
        my_dict[self.w18.labelBetonalatoupie.text()] = self.w18.Betonalatoupie.text()
        my_dict[self.w18.labelchargeuses.text()] = self.w18.chargeuses.text()
        my_dict[self.w18.labeldatederetourmails.text()] = self.w18.datederetourmails.text()
        my_dict[self.w18.labelIDEPOT.text()] = self.w18.IDEPOT.text()
        my_dict[self.w18.labelDecoupeBois.text()] = self.w18.DecoupeBois.text()
        my_dict[self.w18.labelTestRenfortequipe.text()] = self.w18.TestRenfortequipe.text()
        my_dict[self.w18.labelTestsurMesureplacards.text()] = self.w18.TestsurMesureplacards.text()
        my_dict[self.w18.labelTestsurMesureMenuiserie.text()] = self.w18.TestsurMesureMenuiserie.text()
        my_dict[self.w18.labelTestTelephoniesurIP.text()] = self.w18.TestTelephoniesurIP.text()
        my_dict[self.w18.labelTestRETENCYTrackingClients.text()] = self.w18.TestRETENCYTrackingClients.text()
        my_dict[self.w18.labelPresentoiraDalles.text()] = self.w18.PresentoiraDalles.text()
        my_dict[self.w18.labelShowroomSalledeBains.text()] = self.w18.ShowroomSalledeBains.text()
        my_dict[self.w18.labelShowroomSalledeBainsDate.text()] = self.w18.ShowroomSalledeBainsDate.text()
        my_dict[self.w18.labelErgosquelette.text()] = self.w18.Ergosquelette.text()
        my_dict[self.w18.labelTranspaletteciseaux.text()] = self.w18.Transpaletteciseaux.text()
        my_dict[self.w18.labelPlaquesteflonees.text()] = self.w18.Plaquesteflonees.text()
        my_dict[self.w18.labelCagesPalettes.text()] = self.w18.CagesPalettes.text()
        my_dict[self.w19.labelLesclientspenetrentdanslazonedesbureauxpouraccederauxtoilettes.text()] = self.w19.Lesclientspenetrentdanslazonedesbureauxpouraccederauxtoilettes.text()
        my_dict[self.w19.labelSocietedeNettoyage.text()] = self.w19.SocietedeNettoyage.text()
        my_dict[self.w19.labelFenetresaletagedesbureaux.text()] = self.w19.Fenetresaletagedesbureaux.text()
        my_dict[self.w19.labelComptageClients.text()] = self.w19.ComptageClients.text()
        my_dict[self.w19.labelContexteComptageTCUENTO.text()] = self.w19.ContexteComptageTCUENTO.text()
        my_dict[self.w20.labelNdegTVA.text()] = self.w20.NdegTVA.text()
        my_dict[self.w20.labelNdegCommercant.text()] = self.w20.NdegCommercant.text()
        my_dict[self.w20.labelCodeAbonneTransax.text()] = self.w20.CodeAbonneTransax.text()
        my_dict[self.w20.labelCodeVendeurCACF.text()] = self.w20.CodeVendeurCACF.text()
        my_dict[self.w20.labelCodeIPCACF.text()] = self.w20.CodeIPCACF.text()
        my_dict[self.w20.labelNdegCommercantdrive.text()] = self.w20.NdegCommercantdrive.text()
        my_dict[self.w20.labelNdegCommercantPayementsanscontact.text()] = self.w20.NdegCommercantPayementsanscontact.text()
        my_dict[self.w20.labelCodeIBANBNP.text()] = self.w20.CodeIBANBNP.text()
        my_dict[self.w21.labelCompteIdentifiant.text()] = self.w21.CompteIdentifiant.text()
        my_dict[self.w21.labelMotdepasse.text()] = self.w21.Motdepasse.text()
        my_dict[self.w22.labelTauxdAT2015.text()] = self.w22.TauxdAT2015.text()
        my_dict[self.w22.labelTauxdAT2016.text()] = self.w22.TauxdAT2016.text()
        my_dict[self.w22.labelTauxdAT2017.text()] = self.w22.TauxdAT2017.text()
        my_dict[self.w22.labelTauxdAT2018.text()] = self.w22.TauxdAT2018.text()
        my_dict[self.w22.labelTauxdAT2019.text()] = self.w22.TauxdAT2019.text()
        my_dict[self.w22.labelTauxdAT2020.text()] = self.w22.TauxdAT2020.text()
        my_dict[self.w22.labelTauxdAT2021.text()] = self.w22.TauxdAT2021.text()
        
        """
        print (my_dict)
        print (len(my_dict))
        """
        s = pd.Series(my_dict)
        newRow = s.to_frame().T

        #Concaténation avec la sheet originale
        newDF = pd.concat([self.sheet,newRow], axis=0)
        print (newDF)

        main.sheet = newDF
        main.chargerModif()
        self.closeall()

       

    def closeEvent(self, event):
        print("je fais le close event")
        self.w=''
        if self.w:
            self.w.close()

    def closeall(self):
        print("prout pd")
        for child in self.findChildren(QWidget):
            child.close()
        # Fermer cette fenêtre
        self.close()

##DONE1
class listeDepot(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle ("Liste dépôt")
        self.layout = QGridLayout()
        self.titre = QLabel("Liste Dépot")

        self.labelcodeBRICO = QLabel("Code BRICO")
        self.codeBRICO = QLineEdit()
        self.codeBRICO.setPlaceholderText("Code BRICO")
        self.labelcodeEASIER = QLabel("Code EASIER")
        self.codeEASIER = QLineEdit()
        self.codeEASIER.setPlaceholderText("Code EASIER")
        self.labeldepot = QLabel("Dépôt")
        self.depot = QLineEdit()
        self.depot.setPlaceholderText("Dépôt")
        self.labelregionAdministrative = QLabel("Région administrative")
        self.regionAdministrative = QLineEdit() 
        self.regionAdministrative.setPlaceholderText("Région administrative")
        self.labelregion2022 = QLabel("Région 2022")
        self.region2022 = QLineEdit()
        self.region2022.setPlaceholderText("Région 2022")
        self.labelstatut = QLabel("Statut")
        self.statut = QLineEdit()
        self.statut.setPlaceholderText("Statut")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.labelcodeBRICO, 1, 0)
        self.layout.addWidget(self.codeBRICO, 1, 1, 1, 2)
        self.layout.addWidget(self.labelcodeEASIER, 2, 0)
        self.layout.addWidget(self.codeEASIER, 2, 1, 1, 2)
        self.layout.addWidget(self.labeldepot, 3, 0)
        self.layout.addWidget(self.depot, 3, 1, 1, 2)
        self.layout.addWidget(self.labelregionAdministrative, 4, 0)
        self.layout.addWidget(self.regionAdministrative, 4, 1, 1, 2)
        self.layout.addWidget(self.labelregion2022, 5, 0)
        self.layout.addWidget(self.region2022, 5, 1, 1, 2)
        self.layout.addWidget(self.labelstatut, 6, 0)
        self.layout.addWidget(self.statut, 6, 1, 1, 2)
        
        self.setLayout(self.layout)

##DONE2
class donneesSociales(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Données sociales")
        self.layout = QGridLayout()
        self.titre = QLabel("Données sociales")

        self.labelDirecteurRegional = QLabel("Directeur Régional")
        self.DirecteurRegional = QLineEdit()
        self.DirecteurRegional.setPlaceholderText("Directeur Régional")
        self.labelDirecteurRegionaltelephone = QLabel("Directeur Régional téléphone")
        self.DirecteurRegionaltelephone = QLineEdit()
        self.DirecteurRegionaltelephone.setPlaceholderText("Directeur Régional téléphone")
        self.labelDirecteurdepot = QLabel("Directeur dépôt")
        self.Directeurdepot = QLineEdit()
        self.Directeurdepot.setPlaceholderText("Directeur dépôt")
        self.labelDirecteurdepotmail = QLabel("Directeur dépôt mail")
        self.Directeurdepotmail = QLineEdit()
        self.Directeurdepotmail.setPlaceholderText("Directeur dépôt mail")
        self.labelDirecteurdepottelephone = QLabel("Directeur dépôt téléphone")
        self.Directeurdepottelephone = QLineEdit()
        self.Directeurdepottelephone.setPlaceholderText("Directeur dépôt téléphone")
        self.labelDatedeprisedefonction = QLabel("Date de prise de fonction")
        self.Datedeprisedefonction = QLineEdit()
        self.Datedeprisedefonction.setPlaceholderText("Date de prise de fonction")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.labelDirecteurRegional, 1, 0)
        self.layout.addWidget(self.DirecteurRegional, 1, 1, 1, 2)
        self.layout.addWidget(self.labelDirecteurRegionaltelephone, 2, 0)
        self.layout.addWidget(self.DirecteurRegionaltelephone, 2, 1, 1, 2)
        self.layout.addWidget(self.labelDirecteurdepot, 3, 0)
        self.layout.addWidget(self.Directeurdepot, 3, 1, 1, 2)
        self.layout.addWidget(self.labelDirecteurdepotmail, 4, 0)
        self.layout.addWidget(self.Directeurdepotmail, 4, 1, 1, 2)
        self.layout.addWidget(self.labelDirecteurdepottelephone, 5, 0)
        self.layout.addWidget(self.Directeurdepottelephone, 5, 1, 1, 2)
        self.layout.addWidget(self.labelDatedeprisedefonction, 6, 0)
        self.layout.addWidget(self.Datedeprisedefonction, 6, 1, 1, 2)

        self.setLayout(self.layout)

##DONE3
class secteurSecurite(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur sécurité")
        self.layout = QGridLayout()
        self.titre = QLabel("Secteur sécurité")

        self.labelChefSecurite = QLabel("Chef Sécurité")
        self.ChefSecurite = QLineEdit()
        self.ChefSecurite.setPlaceholderText("Chef Sécurité")
        self.labelChefSecuritemail = QLabel("Chef Sécurité mail")
        self.ChefSecuritemail = QLineEdit()
        self.ChefSecuritemail.setPlaceholderText("Chef Sécurité mail")
        self.labelChefSecuriteTelephone = QLabel("Chef Sécurité Téléphone")
        self.ChefSecuriteTelephone = QLineEdit()
        self.ChefSecuriteTelephone.setPlaceholderText("Chef Sécurité Téléphone")
        self.labelChefSecuriteStatut = QLabel("Chef Sécurité Statut")
        self.ChefSecuriteStatut = QLineEdit()
        self.ChefSecuriteStatut.setPlaceholderText("Chef Sécurité Statut")
        self.labelDMfileSecurite = QLabel("DM file Sécurité")
        self.DMfileSecurite = QLineEdit()
        self.DMfileSecurite.setPlaceholderText("DM file Sécurité")
        self.labelCSReferentSecurite = QLabel("CS Référent Sécurité")
        self.CSReferentSecurite = QLineEdit()
        self.CSReferentSecurite.setPlaceholderText("CS Référent Sécurité")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelChefSecurite, 1, 0)
        self.layout.addWidget(self.ChefSecurite, 1, 1, 1, 2)
        self.layout.addWidget(self.labelChefSecuritemail, 2, 0)
        self.layout.addWidget(self.ChefSecuritemail, 2, 1, 1, 2)
        self.layout.addWidget(self.labelChefSecuriteTelephone, 3, 0)
        self.layout.addWidget(self.ChefSecuriteTelephone, 3, 1, 1, 2)
        self.layout.addWidget(self.labelChefSecuriteStatut, 4, 0)
        self.layout.addWidget(self.ChefSecuriteStatut, 4, 1, 1, 2)
        self.layout.addWidget(self.labelDMfileSecurite, 5, 0)
        self.layout.addWidget(self.DMfileSecurite, 5, 1, 1, 2)
        self.layout.addWidget(self.labelCSReferentSecurite, 6, 0)
        self.layout.addWidget(self.CSReferentSecurite, 6, 1, 1, 2)

        self.setLayout(self.layout)

##DONE 4
class secteurLogistique(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur logistique")
        self.layout = QGridLayout()
        self.titre = QLabel("Secteur logistique")
        
        self.labelChefLogistique = QLabel("Chef Logistique")
        self.ChefLogistique = QLineEdit()
        self.ChefLogistique.setPlaceholderText("Chef Logistique")
        self.labelChefLogistiquemail = QLabel("Chef Logistique mail")
        self.ChefLogistiquemail = QLineEdit()
        self.ChefLogistiquemail.setPlaceholderText("Chef Logistique mail")
        self.labelChefLogistiqueTelephone = QLabel("Chef Logistique Téléphone")
        self.ChefLogistiqueTelephone = QLineEdit()
        self.ChefLogistiqueTelephone.setPlaceholderText("Chef Logistique Téléphone")
        self.labelChefLogistiqueStatut = QLabel("Chef Logistique Statut ")
        self.ChefLogistiqueStatut = QLineEdit()
        self.ChefLogistiqueStatut.setPlaceholderText("Chef Logistique Statut ")
        self.labelDMfileLogistique = QLabel("DM file Logistique")
        self.DMfileLogistique = QLineEdit()
        self.DMfileLogistique.setPlaceholderText("DM file Logistique")
        self.labelCSReferentLogistique = QLabel("CS Référent Logistique")
        self.CSReferentLogistique = QLineEdit()
        self.CSReferentLogistique.setPlaceholderText("CS Référent Logistique")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelChefLogistique, 1, 0)
        self.layout.addWidget(self.ChefLogistique, 1, 1, 1, 2)
        self.layout.addWidget(self.labelChefLogistiquemail, 2, 0)
        self.layout.addWidget(self.ChefLogistiquemail, 2, 1, 1, 2)
        self.layout.addWidget(self.labelChefLogistiqueTelephone, 3, 0)
        self.layout.addWidget(self.ChefLogistiqueTelephone, 3, 1, 1, 2)
        self.layout.addWidget(self.labelChefLogistiqueStatut, 4, 0)
        self.layout.addWidget(self.ChefLogistiqueStatut, 4, 1, 1, 2)
        self.layout.addWidget(self.labelDMfileLogistique, 5, 0)
        self.layout.addWidget(self.DMfileLogistique, 5, 1, 1, 2)
        self.layout.addWidget(self.labelCSReferentLogistique, 6, 0)
        self.layout.addWidget(self.CSReferentLogistique, 6, 1, 1, 2)

        self.setLayout(self.layout)

##DONE 5
class secteurAmenagement(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur aménagement")
        self.layout = QGridLayout()
        self.titre = QLabel("Secteur aménagement")

        self.labelChefAmenagement = QLabel("Chef Aménagement")
        self.ChefAmenagement = QLineEdit()
        self.ChefAmenagement.setPlaceholderText("Chef Aménagement")
        self.labelChefAmenagementmail = QLabel("Chef Aménagement mail")
        self.ChefAmenagementmail = QLineEdit()
        self.ChefAmenagementmail.setPlaceholderText("Chef Aménagement mail")
        self.labelChefAmenagementTelephone = QLabel("Chef Aménagement Téléphone")
        self.ChefAmenagementTelephone = QLineEdit()
        self.ChefAmenagementTelephone.setPlaceholderText("Chef Aménagement Téléphone")
        self.labelDMfileAmenagement = QLabel("DM file Aménagement")
        self.DMfileAmenagement = QLineEdit()
        self.DMfileAmenagement.setPlaceholderText("DM file Aménagement")
        self.labelCSReferentAmenagement = QLabel("CS Référent Aménagement")
        self.CSReferentAmenagement = QLineEdit()
        self.CSReferentAmenagement.setPlaceholderText("CS Référent Aménagement")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelChefAmenagement, 1, 0)
        self.layout.addWidget(self.ChefAmenagement, 1, 1, 1, 2)
        self.layout.addWidget(self.labelChefAmenagementmail, 2, 0)
        self.layout.addWidget(self.ChefAmenagementmail, 2, 1, 1, 2)
        self.layout.addWidget(self.labelChefAmenagementTelephone, 3, 0)
        self.layout.addWidget(self.ChefAmenagementTelephone, 3, 1, 1, 2)
        self.layout.addWidget(self.labelDMfileAmenagement, 4, 0)
        self.layout.addWidget(self.DMfileAmenagement, 4, 1, 1, 2)
        self.layout.addWidget(self.labelCSReferentAmenagement, 5, 0)
        self.layout.addWidget(self.CSReferentAmenagement, 5, 1, 1, 2)

        self.setLayout(self.layout)

##DONE6
class secteurConstruction(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur construction")
        self.layout = QGridLayout()
        self.titre = QLabel("Secteur construction")

        self.labelChefConstruction = QLabel("Chef Construction")
        self.ChefConstruction = QLineEdit()
        self.ChefConstruction.setPlaceholderText("Chef Construction")
        self.labelChefConstructionmail = QLabel("Chef Construction mail")
        self.ChefConstructionmail = QLineEdit()
        self.ChefConstructionmail.setPlaceholderText("Chef Construction mail")
        self.labelChefConstructionTelephone = QLabel("Chef Construction Téléphone")
        self.ChefConstructionTelephone = QLineEdit()
        self.ChefConstructionTelephone.setPlaceholderText("Chef Construction Téléphone")
        self.labelDMfileConstruction = QLabel("DM file Construction")
        self.DMfileConstruction = QLineEdit()
        self.DMfileConstruction.setPlaceholderText("DM file Construction")
        self.labelCSReferentConstruction = QLabel("CS Référent Construction")
        self.CSReferentConstruction = QLineEdit()
        self.CSReferentConstruction.setPlaceholderText("CS Référent Construction")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelChefConstruction, 1, 0)
        self.layout.addWidget(self.ChefConstruction, 1, 1, 1, 2)
        self.layout.addWidget(self.labelChefConstructionmail, 2, 0)
        self.layout.addWidget(self.ChefConstructionmail, 2, 1, 1, 2)
        self.layout.addWidget(self.labelChefConstructionTelephone, 3, 0)
        self.layout.addWidget(self.ChefConstructionTelephone, 3, 1, 1, 2)
        self.layout.addWidget(self.labelDMfileConstruction, 4, 0)
        self.layout.addWidget(self.DMfileConstruction, 4, 1, 1, 2)
        self.layout.addWidget(self.labelCSReferentConstruction, 5, 0)
        self.layout.addWidget(self.CSReferentConstruction, 5, 1, 1, 2)

        self.setLayout(self.layout)

##DONE7
class secteurTechnique(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur technique")
        self.layout = QGridLayout()
        self.titre = QLabel("Secteur technique")

        self.labelChefTechnique = QLabel("Chef Technique")
        self.ChefTechnique = QLineEdit()
        self.ChefTechnique.setPlaceholderText("Chef Technique")
        self.labelChefTechniquemail = QLabel("Chef Technique mail")
        self.ChefTechniquemail = QLineEdit()
        self.ChefTechniquemail.setPlaceholderText("Chef Technique mail")
        self.labelChefTechniqueTelephone = QLabel("Chef Technique Téléphone")
        self.ChefTechniqueTelephone = QLineEdit()
        self.ChefTechniqueTelephone.setPlaceholderText("Chef Technique Téléphone")
        self.labelDMfileTechnique = QLabel("DM file Technique")
        self.DMfileTechnique = QLineEdit()
        self.DMfileTechnique.setPlaceholderText("DM file Technique")
        self.labelCSReferentTechnique = QLabel("CS Référent Technique")
        self.CSReferentTechnique = QLineEdit()
        self.CSReferentTechnique.setPlaceholderText("CS Référent Technique")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelChefTechnique, 1, 0)
        self.layout.addWidget(self.ChefTechnique, 1, 1, 1, 2)
        self.layout.addWidget(self.labelChefTechniquemail, 2, 0)
        self.layout.addWidget(self.ChefTechniquemail, 2, 1, 1, 2)
        self.layout.addWidget(self.labelChefTechniqueTelephone, 3, 0)
        self.layout.addWidget(self.ChefTechniqueTelephone, 3, 1, 1, 2)
        self.layout.addWidget(self.labelDMfileTechnique, 4, 0)
        self.layout.addWidget(self.DMfileTechnique, 4, 1, 1, 2)
        self.layout.addWidget(self.labelCSReferentTechnique, 5, 0)
        self.layout.addWidget(self.CSReferentTechnique, 5, 1, 1, 2)

        self.setLayout(self.layout)

##DONE8
class secteurAdministratif(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur administratif")
        self.layout = QGridLayout()
        self.titre = QLabel("Secteur administratif")

        self.labelChefAdministratif = QLabel("Chef Administratif")
        self.ChefAdministratif = QLineEdit()
        self.ChefAdministratif.setPlaceholderText("Chef Administratif")
        self.labelChefAdministratifmail = QLabel("Chef Administratif mail")
        self.ChefAdministratifmail = QLineEdit()
        self.ChefAdministratifmail.setPlaceholderText("Chef Administratif mail")
        self.labelChefAdministratifTelephone = QLabel("Chef Administratif Téléphone")
        self.ChefAdministratifTelephone = QLineEdit()
        self.ChefAdministratifTelephone.setPlaceholderText("Chef Administratif Téléphone")
        self.labelDMfileAdministratif = QLabel("DM file Administratif")
        self.DMfileAdministratif = QLineEdit()
        self.DMfileAdministratif.setPlaceholderText("DM file Administratif")
        self.labelCSReferentAdministratif = QLabel("CS Référent Administratif")
        self.CSReferentAdministratif = QLineEdit()
        self.CSReferentAdministratif.setPlaceholderText("CS Référent Administratif")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelChefAdministratif, 1, 0)
        self.layout.addWidget(self.ChefAdministratif, 1, 1, 1, 2)
        self.layout.addWidget(self.labelChefAdministratifmail, 2, 0)
        self.layout.addWidget(self.ChefAdministratifmail, 2, 1, 1, 2)
        self.layout.addWidget(self.labelChefAdministratifTelephone, 3, 0)
        self.layout.addWidget(self.ChefAdministratifTelephone, 3, 1, 1, 2)
        self.layout.addWidget(self.labelDMfileAdministratif, 4, 0)
        self.layout.addWidget(self.DMfileAdministratif, 4, 1, 1, 2)
        self.layout.addWidget(self.labelCSReferentAdministratif, 5, 0)
        self.layout.addWidget(self.CSReferentAdministratif, 5, 1, 1, 2)

        self.setLayout(self.layout)

##DONE9
class secteurCaisse(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Secteur caisse")
        self.layout = QGridLayout()
        self.titre = QLabel("Secteur caisse")

        self.labelChefCaisse = QLabel("Chef Caisse")
        self.ChefCaisse = QLineEdit()
        self.ChefCaisse.setPlaceholderText("Chef Caisse")
        self.labelChefCaissemail = QLabel("Chef Caisse mail")
        self.ChefCaissemail = QLineEdit()
        self.ChefCaissemail.setPlaceholderText("Chef Caisse mail")
        self.labelChefCaisseTelephone = QLabel("Chef Caisse Téléphone")
        self.ChefCaisseTelephone = QLineEdit()
        self.ChefCaisseTelephone.setPlaceholderText("Chef Caisse Téléphone")
        self.labelDMfileCaisse = QLabel("DM file Caisse")
        self.DMfileCaisse = QLineEdit()
        self.DMfileCaisse.setPlaceholderText("DM file Caisse")
        self.labelCSReferentCaisse = QLabel("CS Référent Caisse")
        self.CSReferentCaisse = QLineEdit()
        self.CSReferentCaisse.setPlaceholderText("CS Référent Caisse")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelChefCaisse, 1, 0)
        self.layout.addWidget(self.ChefCaisse, 1, 1, 1, 2)
        self.layout.addWidget(self.labelChefCaissemail, 2, 0)
        self.layout.addWidget(self.ChefCaissemail, 2, 1, 1, 2)
        self.layout.addWidget(self.labelChefCaisseTelephone, 3, 0)
        self.layout.addWidget(self.ChefCaisseTelephone, 3, 1, 1, 2)
        self.layout.addWidget(self.labelDMfileCaisse, 4, 0)
        self.layout.addWidget(self.DMfileCaisse, 4, 1, 1, 2)
        self.layout.addWidget(self.labelCSReferentCaisse, 5, 0)
        self.layout.addWidget(self.CSReferentCaisse, 5, 1, 1, 2)

        self.setLayout(self.layout)

##DONE10
class RH(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RH")
        self.layout = QGridLayout()
        self.titre = QLabel("RH")

        self.labelResponsableRH = QLabel("Responsable RH ")
        self.ResponsableRH = QLineEdit()
        self.ResponsableRH.setPlaceholderText("Responsable RH ")
        self.labelResponsableRHTelephone = QLabel("Responsable RH Téléphone")
        self.ResponsableRHTelephone = QLineEdit()
        self.ResponsableRHTelephone.setPlaceholderText("Responsable RH Téléphone")
        self.labelResponsableCDG = QLabel("Responsable CDG")
        self.ResponsableCDG = QLineEdit()
        self.ResponsableCDG.setPlaceholderText("Responsable CDG")
        self.labelResponsableCDGTelephone = QLabel("Responsable CDG Téléphone")
        self.ResponsableCDGTelephone = QLineEdit()
        self.ResponsableCDGTelephone.setPlaceholderText("Responsable CDG Téléphone")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelResponsableRH, 1, 0)
        self.layout.addWidget(self.ResponsableRH, 1, 1, 1, 2)
        self.layout.addWidget(self.labelResponsableRHTelephone, 2, 0)
        self.layout.addWidget(self.ResponsableRHTelephone, 2, 1, 1, 2)
        self.layout.addWidget(self.labelResponsableCDG, 3, 0)
        self.layout.addWidget(self.ResponsableCDG, 3, 1, 1, 2)
        self.layout.addWidget(self.labelResponsableCDGTelephone, 4, 0)
        self.layout.addWidget(self.ResponsableCDGTelephone, 4, 1, 1, 2)

        self.setLayout(self.layout)

##DONE: ATTENTION PRESENCE DE ' , de () et de +SIRET (remplacé par SIRET) 11
class donneesDepot(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Données dépôt")
        self.setFixedSize(520,400)
        
        #ScrollArea
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFixedSize(500,380)

        #Widget
        self.widget = QWidget()
        self.layout = QGridLayout()
        self.titre = QLabel("Données dépôt")

        self.labelAdresse = QLabel("Adresse")
        self.Adresse = QLineEdit()
        self.Adresse.setPlaceholderText("Adresse")
        self.labelVille = QLabel("Ville")
        self.Ville = QLineEdit()
        self.Ville.setPlaceholderText("Ville")
        self.labelDepartement = QLabel("Département")
        self.Departement = QLineEdit()
        self.Departement.setPlaceholderText("Département")
        self.labelTELEPHONESTANDARD = QLabel("TELEPHONE STANDARD")
        self.TELEPHONESTANDARD = QLineEdit()
        self.TELEPHONESTANDARD.setPlaceholderText("TELEPHONE STANDARD")
        self.labelLatitudeendegresdecimaux = QLabel("Latitude (en degrés décimaux)")
        self.Latitudeendegresdecimaux = QLineEdit()
        self.Latitudeendegresdecimaux.setPlaceholderText("Latitude (en degrés décimaux)")
        self.labelLongitudeendegresdecimaux = QLabel("Longitude (en degrés décimaux)")
        self.Longitudeendegresdecimaux = QLineEdit()
        self.Longitudeendegresdecimaux.setPlaceholderText("Longitude (en degrés décimaux)")
        self.labelNdegSIREN = QLabel("N° SIREN")
        self.NdegSIREN = QLineEdit()
        self.NdegSIREN.setPlaceholderText("N° SIREN")
        self.labelSIRET = QLabel("+ SIRET")
        self.SIRET = QLineEdit()
        self.SIRET.setPlaceholderText("+ SIRET")
        self.labelDatedOuverture = QLabel("Date d'Ouverture")
        self.DatedOuverture = QLineEdit()
        self.DatedOuverture.setPlaceholderText("Date d'Ouverture")
        self.labelAnciennetedudepot = QLabel("Ancienneté du dépôt")
        self.Anciennetedudepot = QLineEdit()
        self.Anciennetedudepot.setPlaceholderText("Ancienneté du dépôt")
        self.labelHorairedouverturejours1 = QLabel("Horaire d'ouverture jours 1")
        self.Horairedouverturejours1 = QLineEdit()
        self.Horairedouverturejours1.setPlaceholderText("Horaire d'ouverture jours 1")
        self.labelHorairedouverturehoraires1 = QLabel("Horaire d'ouverture horaires 1")
        self.Horairedouverturehoraires1 = QLineEdit()
        self.Horairedouverturehoraires1.setPlaceholderText("Horaire d'ouverture horaires 1")
        self.labelHorairedouverturejours2 = QLabel("Horaire d'ouverture jours 2")
        self.Horairedouverturejours2 = QLineEdit()
        self.Horairedouverturejours2.setPlaceholderText("Horaire d'ouverture jours 2")
        self.labelHorairedouverturehoraires2 = QLabel("Horaire d'ouverture horaires 2")
        self.Horairedouverturehoraires2 = QLineEdit()
        self.Horairedouverturehoraires2.setPlaceholderText("Horaire d'ouverture horaires 2")
        self.labelHorairedouverturejours3 = QLabel("Horaire d'ouverture jours 3")
        self.Horairedouverturejours3 = QLineEdit()
        self.Horairedouverturejours3.setPlaceholderText("Horaire d'ouverture jours 3")
        self.labelHorairedouverturehoraires3 = QLabel("Horaire d'ouverture horaires 3")
        self.Horairedouverturehoraires3 = QLineEdit()
        self.Horairedouverturehoraires3.setPlaceholderText("Horaire d'ouverture horaires 3")
        self.labelAmplitudeHoraire = QLabel("Amplitude Horaire")
        self.AmplitudeHoraire = QLineEdit()
        self.AmplitudeHoraire.setPlaceholderText("Amplitude Horaire")
        self.labelCATTC2021 = QLabel("CA TTC 2021")
        self.CATTC2021 = QLineEdit()
        self.CATTC2021.setPlaceholderText("CA TTC 2021")
        self.labelBASSIN2021 = QLabel("BASSIN 2021")
        self.BASSIN2021 = QLineEdit()
        self.BASSIN2021.setPlaceholderText("BASSIN 2021")
        self.labelCluster2021 = QLabel("Cluster 2021")
        self.Cluster2021 = QLineEdit()
        self.Cluster2021.setPlaceholderText("Cluster 2021")
        self.labelPassagesCaisses2021 = QLabel("Passages Caisses 2021")
        self.PassagesCaisses2021 = QLineEdit()
        self.PassagesCaisses2021.setPlaceholderText("Passages Caisses 2021")
        self.labelTauxdeDemarque2021 = QLabel("Taux de Démarque 2021")
        self.TauxdeDemarque2021 = QLineEdit()
        self.TauxdeDemarque2021.setPlaceholderText("Taux de Démarque 2021")
        self.labelETPCDI2021 = QLabel("ETP CDI 2021")
        self.ETPCDI2021 = QLineEdit()
        self.ETPCDI2021.setPlaceholderText("ETP CDI 2021")
        self.labelETPCDD2021 = QLabel("ETP CDD 2021")
        self.ETPCDD2021 = QLineEdit()
        self.ETPCDD2021.setPlaceholderText("ETP CDD 2021")
        self.labelETPINTERIM2021 = QLabel("ETP INTERIM 2021")
        self.ETPINTERIM2021 = QLineEdit()
        self.ETPINTERIM2021.setPlaceholderText("ETP INTERIM 2021")
        self.labelETPGLOBAL2021 = QLabel("ETP GLOBAL 2021")
        self.ETPGLOBAL2021 = QLineEdit()
        self.ETPGLOBAL2021.setPlaceholderText("ETP GLOBAL 2021")
        self.labelFormatdepot = QLabel("Format dépôt")
        self.Formatdepot = QLineEdit()
        self.Formatdepot.setPlaceholderText("Format dépôt")
        self.labelNombredesites = QLabel("Nombre de sites")
        self.Nombredesites = QLineEdit()
        self.Nombredesites.setPlaceholderText("Nombre de sites")
        self.labelZoneScolaire = QLabel("Zone Scolaire")
        self.ZoneScolaire = QLineEdit()
        self.ZoneScolaire.setPlaceholderText("Zone Scolaire")
        self.labelZonedeDroitSocial = QLabel("Zone de Droit Social")
        self.ZonedeDroitSocial = QLineEdit()
        self.ZonedeDroitSocial.setPlaceholderText("Zone de Droit Social")
        self.labelEntrepotLegacy = QLabel("Entrepôt Legacy")
        self.EntrepotLegacy = QLineEdit()
        self.EntrepotLegacy.setPlaceholderText("Entrepôt Legacy")
        self.labelEntrepotEasier = QLabel("Entrepôt Easier")
        self.EntrepotEasier = QLineEdit()
        self.EntrepotEasier.setPlaceholderText("Entrepôt Easier")
        self.labelPlateformeLegacy = QLabel("Plateforme Legacy")
        self.PlateformeLegacy = QLineEdit()
        self.PlateformeLegacy.setPlaceholderText("Plateforme Legacy")
        self.labelPlateformeEasier = QLabel("Plateforme Easier")
        self.PlateformeEasier = QLineEdit()
        self.PlateformeEasier.setPlaceholderText("Plateforme Easier")
        self.labelPlateformeArrivage = QLabel("Plateforme Arrivage")
        self.PlateformeArrivage = QLineEdit()
        self.PlateformeArrivage.setPlaceholderText("Plateforme Arrivage")
        self.labelPlateformeArrivageECC = QLabel("Plateforme Arrivage ECC")
        self.PlateformeArrivageECC = QLineEdit()
        self.PlateformeArrivageECC.setPlaceholderText("Plateforme Arrivage ECC")
        self.labelPlateformederattachementGammeetArrivagesapresouvertures = QLabel("Plateforme de rattachement (Gamme et Arrivages) après ouvertures")
        self.PlateformederattachementGammeetArrivagesapresouvertures = QLineEdit()
        self.PlateformederattachementGammeetArrivagesapresouvertures.setPlaceholderText("Plateforme de rattachement (Gamme et Arrivages) après ouvertures")
        self.labelGLNPlateformeGamme = QLabel("GLN Plateforme Gamme")
        self.GLNPlateformeGamme = QLineEdit()
        self.GLNPlateformeGamme.setPlaceholderText("GLN Plateforme Gamme")
        self.labelCodeplateforme = QLabel("Code plateforme")
        self.Codeplateforme = QLineEdit()
        self.Codeplateforme.setPlaceholderText("Code plateforme")
        self.labelEntrepotdeporte = QLabel("Entrepôt déporté")
        self.Entrepotdeporte = QLineEdit()
        self.Entrepotdeporte.setPlaceholderText("Entrepôt déporté")
        self.labelOrigineduDepot = QLabel("Origine du Dépôt")
        self.OrigineduDepot = QLineEdit()
        self.OrigineduDepot.setPlaceholderText("Origine du Dépôt")
        self.labelSituationduDepot = QLabel("Situation du Dépôt")
        self.SituationduDepot = QLineEdit()
        self.SituationduDepot.setPlaceholderText("Situation du Dépôt")
        self.labelProprietaireouLocataire = QLabel("Propriétaire ou Locataire")
        self.ProprietaireouLocataire = QLineEdit()
        self.ProprietaireouLocataire.setPlaceholderText("Propriétaire ou Locataire")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelAdresse, 1, 0)
        self.layout.addWidget(self.Adresse, 1, 1, 1, 2)
        self.layout.addWidget(self.labelVille, 2, 0)
        self.layout.addWidget(self.Ville, 2, 1, 1, 2)
        self.layout.addWidget(self.labelDepartement, 3, 0)
        self.layout.addWidget(self.Departement, 3, 1, 1, 2)
        self.layout.addWidget(self.labelTELEPHONESTANDARD, 4, 0)
        self.layout.addWidget(self.TELEPHONESTANDARD, 4, 1, 1, 2)
        self.layout.addWidget(self.labelLatitudeendegresdecimaux, 5, 0)
        self.layout.addWidget(self.Latitudeendegresdecimaux, 5, 1, 1, 2)
        self.layout.addWidget(self.labelLongitudeendegresdecimaux, 6, 0)
        self.layout.addWidget(self.Longitudeendegresdecimaux, 6, 1, 1, 2)
        self.layout.addWidget(self.labelNdegSIREN, 7, 0)
        self.layout.addWidget(self.NdegSIREN, 7, 1, 1, 2)
        self.layout.addWidget(self.labelSIRET, 8, 0)
        self.layout.addWidget(self.SIRET, 8, 1, 1, 2)
        self.layout.addWidget(self.labelDatedOuverture, 9, 0)
        self.layout.addWidget(self.DatedOuverture, 9, 1, 1, 2)
        self.layout.addWidget(self.labelAnciennetedudepot, 10, 0)
        self.layout.addWidget(self.Anciennetedudepot, 10, 1, 1, 2)
        self.layout.addWidget(self.labelHorairedouverturejours1, 11, 0)
        self.layout.addWidget(self.Horairedouverturejours1, 11, 1, 1, 2)
        self.layout.addWidget(self.labelHorairedouverturehoraires1, 12, 0)
        self.layout.addWidget(self.Horairedouverturehoraires1, 12, 1, 1, 2)
        self.layout.addWidget(self.labelHorairedouverturejours2, 13, 0)
        self.layout.addWidget(self.Horairedouverturejours2, 13, 1, 1, 2)
        self.layout.addWidget(self.labelHorairedouverturehoraires2, 14, 0)
        self.layout.addWidget(self.Horairedouverturehoraires2, 14, 1, 1, 2)
        self.layout.addWidget(self.labelHorairedouverturejours3, 15, 0)
        self.layout.addWidget(self.Horairedouverturejours3, 15, 1, 1, 2)
        self.layout.addWidget(self.labelHorairedouverturehoraires3, 16, 0)
        self.layout.addWidget(self.Horairedouverturehoraires3, 16, 1, 1, 2)
        self.layout.addWidget(self.labelAmplitudeHoraire, 17, 0)
        self.layout.addWidget(self.AmplitudeHoraire, 17, 1, 1, 2)
        self.layout.addWidget(self.labelCATTC2021, 18, 0)
        self.layout.addWidget(self.CATTC2021, 18, 1, 1, 2)
        self.layout.addWidget(self.labelBASSIN2021, 19, 0)
        self.layout.addWidget(self.BASSIN2021, 19, 1, 1, 2)
        self.layout.addWidget(self.labelCluster2021, 20, 0)
        self.layout.addWidget(self.Cluster2021, 20, 1, 1, 2)
        self.layout.addWidget(self.labelPassagesCaisses2021, 21, 0)
        self.layout.addWidget(self.PassagesCaisses2021, 21, 1, 1, 2)
        self.layout.addWidget(self.labelTauxdeDemarque2021, 22, 0)
        self.layout.addWidget(self.TauxdeDemarque2021, 22, 1, 1, 2)
        self.layout.addWidget(self.labelETPCDI2021, 23, 0)
        self.layout.addWidget(self.ETPCDI2021, 23, 1, 1, 2)
        self.layout.addWidget(self.labelETPCDD2021, 24, 0)
        self.layout.addWidget(self.ETPCDD2021, 24, 1, 1, 2)
        self.layout.addWidget(self.labelETPINTERIM2021, 25, 0)
        self.layout.addWidget(self.ETPINTERIM2021, 25, 1, 1, 2)
        self.layout.addWidget(self.labelETPGLOBAL2021, 26, 0)
        self.layout.addWidget(self.ETPGLOBAL2021, 26, 1, 1, 2)
        self.layout.addWidget(self.labelFormatdepot, 27, 0)
        self.layout.addWidget(self.Formatdepot, 27, 1, 1, 2)
        self.layout.addWidget(self.labelNombredesites, 28, 0)
        self.layout.addWidget(self.Nombredesites, 28, 1, 1, 2)
        self.layout.addWidget(self.labelZoneScolaire, 29, 0)
        self.layout.addWidget(self.ZoneScolaire, 29, 1, 1, 2)
        self.layout.addWidget(self.labelZonedeDroitSocial, 30, 0)
        self.layout.addWidget(self.ZonedeDroitSocial, 30, 1, 1, 2)
        self.layout.addWidget(self.labelEntrepotLegacy, 31, 0)
        self.layout.addWidget(self.EntrepotLegacy, 31, 1, 1, 2)
        self.layout.addWidget(self.labelEntrepotEasier, 32, 0)
        self.layout.addWidget(self.EntrepotEasier, 32, 1, 1, 2)
        self.layout.addWidget(self.labelPlateformeLegacy, 33, 0)
        self.layout.addWidget(self.PlateformeLegacy, 33, 1, 1, 2)
        self.layout.addWidget(self.labelPlateformeEasier, 34, 0)
        self.layout.addWidget(self.PlateformeEasier, 34, 1, 1, 2)
        self.layout.addWidget(self.labelPlateformeArrivage, 35, 0)
        self.layout.addWidget(self.PlateformeArrivage, 35, 1, 1, 2)
        self.layout.addWidget(self.labelPlateformeArrivageECC, 36, 0)
        self.layout.addWidget(self.PlateformeArrivageECC, 36, 1, 1, 2)
        self.layout.addWidget(self.labelPlateformederattachementGammeetArrivagesapresouvertures, 37, 0)
        self.layout.addWidget(self.PlateformederattachementGammeetArrivagesapresouvertures, 37, 1, 1, 2)
        self.layout.addWidget(self.labelGLNPlateformeGamme, 38, 0)
        self.layout.addWidget(self.GLNPlateformeGamme, 38, 1, 1, 2)
        self.layout.addWidget(self.labelCodeplateforme, 39, 0)
        self.layout.addWidget(self.Codeplateforme, 39, 1, 1, 2)
        self.layout.addWidget(self.labelEntrepotdeporte, 40, 0)
        self.layout.addWidget(self.Entrepotdeporte, 40, 1, 1, 2)
        self.layout.addWidget(self.labelOrigineduDepot, 41, 0)
        self.layout.addWidget(self.OrigineduDepot, 41, 1, 1, 2)
        self.layout.addWidget(self.labelSituationduDepot, 42, 0)
        self.layout.addWidget(self.SituationduDepot, 42, 1, 1, 2)
        self.layout.addWidget(self.labelProprietaireouLocataire, 43, 0)
        self.layout.addWidget(self.ProprietaireouLocataire, 43, 1, 1, 2)

        self.widget.setLayout(self.layout)

        self.scroll_area.setWidget(self.widget)

##DONE: ATTENTION PRESENCE DE TRUC SVI ET DE / 12
class surface(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Surface")

        self.setFixedSize(520, 400)

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFixedSize(500, 380)

        self.widget = QWidget()
        self.layout = QGridLayout()
        self.titre = QLabel("Surface")

        self.labelSurfaceTotaleCDAC = QLabel("Surface Totale CDAC")
        self.SurfaceTotaleCDAC = QLineEdit()
        self.SurfaceTotaleCDAC.setPlaceholderText("Surface Totale CDAC")
        self.labelSurfacedeventeinterieure = QLabel("Surface de vente intérieure")
        self.Surfacedeventeinterieure = QLineEdit()
        self.Surfacedeventeinterieure.setPlaceholderText("Surface de vente intérieure")
        self.labelTypedesurfaceSVI = QLabel("Type de surface SVI")
        self.TypedesurfaceSVI = QLineEdit()
        self.TypedesurfaceSVI.setPlaceholderText("Type de surface SVI")
        self.labelTypologieSVI = QLabel("Typologie SVI")
        self.TypologieSVI = QLineEdit()
        self.TypologieSVI.setPlaceholderText("Typologie SVI")
        self.labelCourMateriaux = QLabel("Cour Matériaux")
        self.CourMateriaux = QLineEdit()
        self.CourMateriaux.setPlaceholderText("Cour Matériaux")
        self.labelDistanceCourMateriauxDepot = QLabel("Distance Cour Matériaux Dépôt")
        self.DistanceCourMateriauxDepot = QLineEdit()
        self.DistanceCourMateriauxDepot.setPlaceholderText("Distance Cour Matériaux Dépôt")
        self.labelSurfaceBati = QLabel("Surface Bâti")
        self.SurfaceBati = QLineEdit()
        self.SurfaceBati.setPlaceholderText("Surface Bâti")
        self.labelTypedesurfaceBati = QLabel("Type de surface Bâti")
        self.TypedesurfaceBati = QLineEdit()
        self.TypedesurfaceBati.setPlaceholderText("Type de surface Bâti")
        self.labelTypologieBati = QLabel("Typologie Bâti")
        self.TypologieBati = QLineEdit()
        self.TypologieBati.setPlaceholderText("Typologie Bâti")
        self.labelSurfacedesmateriauxenCDAC = QLabel("Surface des matériaux en CDAC")
        self.SurfacedesmateriauxenCDAC = QLineEdit()
        self.SurfacedesmateriauxenCDAC.setPlaceholderText("Surface des matériaux en CDAC")
        self.labelSURFACEBatienCDAC = QLabel("SURFACE Bâti en CDAC")
        self.SURFACEBatienCDAC = QLineEdit()
        self.SURFACEBatienCDAC.setPlaceholderText("SURFACE Bâti en CDAC")
        self.labeldontLocaldeventeBatienCDAC = QLabel("dont Local de vente Bâti en CDAC")
        self.dontLocaldeventeBatienCDAC = QLineEdit()
        self.dontLocaldeventeBatienCDAC.setPlaceholderText("dont Local de vente Bâti en CDAC")
        self.labeldontGNBensurfacedevente = QLabel("dont GNB en surface de vente")
        self.dontGNBensurfacedevente = QLineEdit()
        self.dontGNBensurfacedevente.setPlaceholderText("dont GNB en surface de vente")
        self.labeldontBatiCouvertenCDAC = QLabel("dont Bâti Couvert en CDAC")
        self.dontBatiCouvertenCDAC = QLineEdit()
        self.dontBatiCouvertenCDAC.setPlaceholderText("dont Bâti Couvert en CDAC")
        self.labeldontBatiNoncouvertenCDAC = QLabel("dont Bâti Non couvert en CDAC")
        self.dontBatiNoncouvertenCDAC = QLineEdit()
        self.dontBatiNoncouvertenCDAC.setPlaceholderText("dont Bâti Non couvert en CDAC")
        self.labelSurfacedelacourmateriauxhorsCDAC = QLabel("Surface de la cour matériaux hors CDAC")
        self.SurfacedelacourmateriauxhorsCDAC = QLineEdit()
        self.SurfacedelacourmateriauxhorsCDAC.setPlaceholderText("Surface de la cour matériaux hors CDAC")
        self.labeldontBatiCouverthorsCDAC = QLabel("dont Bâti Couvert hors CDAC")
        self.dontBatiCouverthorsCDAC = QLineEdit()
        self.dontBatiCouverthorsCDAC.setPlaceholderText("dont Bâti Couvert hors CDAC")
        self.labeldontBatiNonCouverthorsCDAC = QLabel("dont Bâti Non Couvert hors CDAC")
        self.dontBatiNonCouverthorsCDAC = QLineEdit()
        self.dontBatiNonCouverthorsCDAC.setPlaceholderText("dont Bâti Non Couvert hors CDAC")
        self.labelEmplacementMenuiserie = QLabel("Emplacement Menuiserie")
        self.EmplacementMenuiserie = QLineEdit()
        self.EmplacementMenuiserie.setPlaceholderText("Emplacement Menuiserie")
        self.labelConfigurationMenuiserie = QLabel("Configuration Menuiserie")
        self.ConfigurationMenuiserie = QLineEdit()
        self.ConfigurationMenuiserie.setPlaceholderText("Configuration Menuiserie")
        self.labelDistanceMenuiserieDepot = QLabel("Distance Menuiserie/ Dépôt")
        self.DistanceMenuiserieDepot = QLineEdit()
        self.DistanceMenuiserieDepot.setPlaceholderText("Distance Menuiserie/ Dépôt")
        self.labelShowRoomMenuiserie = QLabel("Show Room Menuiserie")
        self.ShowRoomMenuiserie = QLineEdit()
        self.ShowRoomMenuiserie.setPlaceholderText("Show Room Menuiserie")
        self.labelSurfacedelamenuiserie = QLabel("Surface de la menuiserie")
        self.Surfacedelamenuiserie = QLineEdit()
        self.Surfacedelamenuiserie.setPlaceholderText("Surface de la menuiserie")
        self.labeldontMenuiserieensurfacedevente = QLabel("dont Menuiserie en surface de vente")
        self.dontMenuiserieensurfacedevente = QLineEdit()
        self.dontMenuiserieensurfacedevente.setPlaceholderText("dont Menuiserie en surface de vente")
        self.labeldontMenuiserieenreserve = QLabel("dont Menuiserie en réserve")
        self.dontMenuiserieenreserve = QLineEdit()
        self.dontMenuiserieenreserve.setPlaceholderText("dont Menuiserie en réserve")
        self.labelShowRoomSalledeBains = QLabel("Show Room Salle de Bains")
        self.ShowRoomSalledeBains = QLineEdit()
        self.ShowRoomSalledeBains.setPlaceholderText("Show Room Salle de Bains")
        self.labelSurfacedelareserve = QLabel("Surface de la réserve")
        self.Surfacedelareserve = QLineEdit()
        self.Surfacedelareserve.setPlaceholderText("Surface de la réserve")
        self.labelSurfaceduSas = QLabel("Surface du Sas")
        self.SurfaceduSas = QLineEdit()
        self.SurfaceduSas.setPlaceholderText("Surface du Sas")
        self.labelSurfacedesbureaux = QLabel("Surface des bureaux")
        self.Surfacedesbureaux = QLineEdit()
        self.Surfacedesbureaux.setPlaceholderText("Surface des bureaux")
        self.labelPlacesdeparking = QLabel("Places de parking")
        self.Placesdeparking = QLineEdit()
        self.Placesdeparking.setPlaceholderText("Places de parking")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelSurfaceTotaleCDAC, 1, 0)
        self.layout.addWidget(self.SurfaceTotaleCDAC, 1, 1, 1, 2)
        self.layout.addWidget(self.labelSurfacedeventeinterieure, 2, 0)
        self.layout.addWidget(self.Surfacedeventeinterieure, 2, 1, 1, 2)
        self.layout.addWidget(self.labelTypedesurfaceSVI, 3, 0)
        self.layout.addWidget(self.TypedesurfaceSVI, 3, 1, 1, 2)
        self.layout.addWidget(self.labelTypologieSVI, 4, 0)
        self.layout.addWidget(self.TypologieSVI, 4, 1, 1, 2)
        self.layout.addWidget(self.labelCourMateriaux, 5, 0)
        self.layout.addWidget(self.CourMateriaux, 5, 1, 1, 2)
        self.layout.addWidget(self.labelDistanceCourMateriauxDepot, 6, 0)
        self.layout.addWidget(self.DistanceCourMateriauxDepot, 6, 1, 1, 2)
        self.layout.addWidget(self.labelSurfaceBati, 7, 0)
        self.layout.addWidget(self.SurfaceBati, 7, 1, 1, 2)
        self.layout.addWidget(self.labelTypedesurfaceBati, 8, 0)
        self.layout.addWidget(self.TypedesurfaceBati, 8, 1, 1, 2)
        self.layout.addWidget(self.labelTypologieBati, 9, 0)
        self.layout.addWidget(self.TypologieBati, 9, 1, 1, 2)
        self.layout.addWidget(self.labelSurfacedesmateriauxenCDAC, 10, 0)
        self.layout.addWidget(self.SurfacedesmateriauxenCDAC, 10, 1, 1, 2)
        self.layout.addWidget(self.labelSURFACEBatienCDAC, 11, 0)
        self.layout.addWidget(self.SURFACEBatienCDAC, 11, 1, 1, 2)
        self.layout.addWidget(self.labeldontLocaldeventeBatienCDAC, 12, 0)
        self.layout.addWidget(self.dontLocaldeventeBatienCDAC, 12, 1, 1, 2)
        self.layout.addWidget(self.labeldontGNBensurfacedevente, 13, 0)
        self.layout.addWidget(self.dontGNBensurfacedevente, 13, 1, 1, 2)
        self.layout.addWidget(self.labeldontBatiCouvertenCDAC, 14, 0)
        self.layout.addWidget(self.dontBatiCouvertenCDAC, 14, 1, 1, 2)
        self.layout.addWidget(self.labeldontBatiNoncouvertenCDAC, 15, 0)
        self.layout.addWidget(self.dontBatiNoncouvertenCDAC, 15, 1, 1, 2)
        self.layout.addWidget(self.labelSurfacedelacourmateriauxhorsCDAC, 16, 0)
        self.layout.addWidget(self.SurfacedelacourmateriauxhorsCDAC, 16, 1, 1, 2)
        self.layout.addWidget(self.labeldontBatiCouverthorsCDAC, 17, 0)
        self.layout.addWidget(self.dontBatiCouverthorsCDAC, 17, 1, 1, 2)
        self.layout.addWidget(self.labeldontBatiNonCouverthorsCDAC, 18, 0)
        self.layout.addWidget(self.dontBatiNonCouverthorsCDAC, 18, 1, 1, 2)
        self.layout.addWidget(self.labelEmplacementMenuiserie, 19, 0)
        self.layout.addWidget(self.EmplacementMenuiserie, 19, 1, 1, 2)
        self.layout.addWidget(self.labelConfigurationMenuiserie, 20, 0)
        self.layout.addWidget(self.ConfigurationMenuiserie, 20, 1, 1, 2)
        self.layout.addWidget(self.labelDistanceMenuiserieDepot, 21, 0)
        self.layout.addWidget(self.DistanceMenuiserieDepot, 21, 1, 1, 2)
        self.layout.addWidget(self.labelShowRoomMenuiserie, 22, 0)
        self.layout.addWidget(self.ShowRoomMenuiserie, 22, 1, 1, 2)
        self.layout.addWidget(self.labelSurfacedelamenuiserie, 23, 0)
        self.layout.addWidget(self.Surfacedelamenuiserie, 23, 1, 1, 2)
        self.layout.addWidget(self.labeldontMenuiserieensurfacedevente, 24, 0)
        self.layout.addWidget(self.dontMenuiserieensurfacedevente, 24, 1, 1, 2)
        self.layout.addWidget(self.labeldontMenuiserieenreserve, 25, 0)
        self.layout.addWidget(self.dontMenuiserieenreserve, 25, 1, 1, 2)
        self.layout.addWidget(self.labelShowRoomSalledeBains, 26, 0)
        self.layout.addWidget(self.ShowRoomSalledeBains, 26, 1, 1, 2)
        self.layout.addWidget(self.labelSurfacedelareserve, 27, 0)
        self.layout.addWidget(self.Surfacedelareserve, 27, 1, 1, 2)
        self.layout.addWidget(self.labelSurfaceduSas, 28, 0)
        self.layout.addWidget(self.SurfaceduSas, 28, 1, 1, 2)
        self.layout.addWidget(self.labelSurfacedesbureaux, 29, 0)
        self.layout.addWidget(self.Surfacedesbureaux, 29, 1, 1, 2)
        self.layout.addWidget(self.labelPlacesdeparking, 30, 0)
        self.layout.addWidget(self.Placesdeparking, 30, 1, 1, 2)

        self.widget.setLayout(self.layout)
        self.scroll_area.setWidget(self.widget)

##DONE 13
class agencement(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Agencement")
        self.layout = QGridLayout()
        self.titre = QLabel("Agencement")

        self.labelFournisseurAgencement = QLabel("Fournisseur Agencement")
        self.FournisseurAgencement = QLineEdit()
        self.FournisseurAgencement.setPlaceholderText("Fournisseur Agencement")
        self.labelNbredenginsdemanutention = QLabel("Nbre d'engins de manutention")
        self.Nbredenginsdemanutention = QLineEdit()
        self.Nbredenginsdemanutention.setPlaceholderText("Nbre d'engins de manutention")
        self.labelSystemedeGestiondeFlotte = QLabel("Système de Gestion de Flotte")
        self.SystemedeGestiondeFlotte = QLineEdit()
        self.SystemedeGestiondeFlotte.setPlaceholderText("Système de Gestion de Flotte")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelFournisseurAgencement, 1, 0)
        self.layout.addWidget(self.FournisseurAgencement, 1, 1, 1, 2)
        self.layout.addWidget(self.labelNbredenginsdemanutention, 2, 0)
        self.layout.addWidget(self.Nbredenginsdemanutention, 2, 1, 1, 2)
        self.layout.addWidget(self.labelSystemedeGestiondeFlotte, 3, 0)
        self.layout.addWidget(self.SystemedeGestiondeFlotte, 3, 1, 1, 2)

        self.setLayout(self.layout)

##DONE 14
class caisse(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Caisse")
        self.layout = QGridLayout()
        self.titre = QLabel("Caisse")

        self.labelPCDediesFlexpoint = QLabel("PC Dédiés Flexpoint")
        self.PCDediesFlexpoint = QLineEdit()
        self.PCDediesFlexpoint.setPlaceholderText("PC Dédiés Flexpoint")
        self.labelNbredecaissesACCUEIL = QLabel("Nbre  de caisses ACCUEIL ")
        self.NbredecaissesACCUEIL = QLineEdit()
        self.NbredecaissesACCUEIL.setPlaceholderText("Nbre  de caisses ACCUEIL ")
        self.labelNbredecaissesMAGASIN = QLabel("Nbre de caisses MAGASIN")
        self.NbredecaissesMAGASIN = QLineEdit()
        self.NbredecaissesMAGASIN.setPlaceholderText("Nbre de caisses MAGASIN")
        self.labelNbredeCaissesBATI = QLabel("Nbre de Caisses BATI")
        self.NbredeCaissesBATI = QLineEdit()
        self.NbredeCaissesBATI.setPlaceholderText("Nbre de Caisses BATI")
        self.labelNbredeCaissesGNB = QLabel("Nbre de Caisses GNB")
        self.NbredeCaissesGNB = QLineEdit()
        self.NbredeCaissesGNB.setPlaceholderText("Nbre de Caisses GNB")
        self.labelNbredeCaissesDEPOUILL = QLabel("Nbre de Caisses DEPOUILL")
        self.NbredeCaissesDEPOUILL = QLineEdit()
        self.NbredeCaissesDEPOUILL.setPlaceholderText("Nbre de Caisses DEPOUILL")
        self.labelTotalCaisses = QLabel("Total Caisses")
        self.TotalCaisses = QLineEdit()
        self.TotalCaisses.setPlaceholderText("Total Caisses")
        self.labelDateRemplacementCaisses = QLabel("Date Remplacement Caisses")
        self.DateRemplacementCaisses = QLineEdit()
        self.DateRemplacementCaisses.setPlaceholderText("Date Remplacement Caisses")
        self.labelNbreSCO = QLabel("Nbre SCO")
        self.NbreSCO = QLineEdit()
        self.NbreSCO.setPlaceholderText("Nbre SCO")
        self.labelModeleSCO = QLabel("Modèle SCO")
        self.ModeleSCO = QLineEdit()
        self.ModeleSCO.setPlaceholderText("Modèle SCO")
        self.labelModeledeTPE = QLabel("Modèle de TPE")
        self.ModeledeTPE = QLineEdit()
        self.ModeledeTPE.setPlaceholderText("Modèle de TPE")
        self.labelDateRemplacementTPE = QLabel("Date Remplacement TPE")
        self.DateRemplacementTPE = QLineEdit()
        self.DateRemplacementTPE.setPlaceholderText("Date Remplacement TPE")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelPCDediesFlexpoint, 1, 0)
        self.layout.addWidget(self.PCDediesFlexpoint, 1, 1, 1, 2)
        self.layout.addWidget(self.labelNbredecaissesACCUEIL, 2, 0)
        self.layout.addWidget(self.NbredecaissesACCUEIL, 2, 1, 1, 2)
        self.layout.addWidget(self.labelNbredecaissesMAGASIN, 3, 0)
        self.layout.addWidget(self.NbredecaissesMAGASIN, 3, 1, 1, 2)
        self.layout.addWidget(self.labelNbredeCaissesBATI, 4, 0)
        self.layout.addWidget(self.NbredeCaissesBATI, 4, 1, 1, 2)
        self.layout.addWidget(self.labelNbredeCaissesGNB, 5, 0)
        self.layout.addWidget(self.NbredeCaissesGNB, 5, 1, 1, 2)
        self.layout.addWidget(self.labelNbredeCaissesDEPOUILL, 6, 0)
        self.layout.addWidget(self.NbredeCaissesDEPOUILL, 6, 1, 1, 2)
        self.layout.addWidget(self.labelTotalCaisses, 7, 0)
        self.layout.addWidget(self.TotalCaisses, 7, 1, 1, 2)
        self.layout.addWidget(self.labelDateRemplacementCaisses, 8, 0)
        self.layout.addWidget(self.DateRemplacementCaisses, 8, 1, 1, 2)
        self.layout.addWidget(self.labelNbreSCO, 9, 0)
        self.layout.addWidget(self.NbreSCO, 9, 1, 1, 2)
        self.layout.addWidget(self.labelModeleSCO, 10, 0)
        self.layout.addWidget(self.ModeleSCO, 10, 1, 1, 2)
        self.layout.addWidget(self.labelModeledeTPE, 11, 0)
        self.layout.addWidget(self.ModeledeTPE, 11, 1, 1, 2)
        self.layout.addWidget(self.labelDateRemplacementTPE, 12, 0)
        self.layout.addWidget(self.DateRemplacementTPE, 12, 1, 1, 2)

        self.setLayout(self.layout)

##DONE 15
class PDA(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDA")
        self.layout = QGridLayout()
        self.titre = QLabel("PDA")

        self.labelDotationPDAMotorolaMC75 = QLabel("Dotation PDA Motorola MC 75")
        self.DotationPDAMotorolaMC75 = QLineEdit()
        self.DotationPDAMotorolaMC75.setPlaceholderText("Dotation PDA Motorola MC 75")
        self.labelPDAMC75Restant012017 = QLabel("PDA MC 75 Restant 01/2017")
        self.PDAMC75Restant012017 = QLineEdit()
        self.PDAMC75Restant012017.setPlaceholderText("PDA MC 75 Restant 01/2017")
        self.labelImprimanteZebra = QLabel("Imprimante Zebra")
        self.ImprimanteZebra = QLineEdit()
        self.ImprimanteZebra.setPlaceholderText("Imprimante Zebra")
        self.labelPDARelevesdeprix = QLabel("PDA Relevés de prix")
        self.PDARelevesdeprix = QLineEdit()
        self.PDARelevesdeprix.setPlaceholderText("PDA Relevés de prix")
        self.labelRescencemntPDA2019 = QLabel("Rescencemnt PDA 2019")
        self.RescencemntPDA2019 = QLineEdit()
        self.RescencemntPDA2019.setPlaceholderText("Rescencemnt PDA 2019")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelDotationPDAMotorolaMC75, 1, 0)
        self.layout.addWidget(self.DotationPDAMotorolaMC75, 1, 1, 1, 2)
        self.layout.addWidget(self.labelPDAMC75Restant012017, 2, 0)
        self.layout.addWidget(self.PDAMC75Restant012017, 2, 1, 1, 2)
        self.layout.addWidget(self.labelImprimanteZebra, 3, 0)
        self.layout.addWidget(self.ImprimanteZebra, 3, 1, 1, 2)
        self.layout.addWidget(self.labelPDARelevesdeprix, 4, 0)
        self.layout.addWidget(self.PDARelevesdeprix, 4, 1, 1, 2)
        self.layout.addWidget(self.labelRescencemntPDA2019, 5, 0)
        self.layout.addWidget(self.RescencemntPDA2019, 5, 1, 1, 2)

        self.setLayout(self.layout)

##DONE16
class menace(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Menace")
        self.layout = QGridLayout()
        self.titre = QLabel("Menace")

        self.labelTauxdevolspour100000habitants = QLabel("Taux de vols pour 100 000 habitants")
        self.Tauxdevolspour100000habitants = QLineEdit()
        self.Tauxdevolspour100000habitants.setPlaceholderText("Taux de vols pour 100 000 habitants")
        self.labelDepotarisquedeVoletnaturedurisque = QLabel("Dépôt à risque de Vol et nature du risque")
        self.DepotarisquedeVoletnaturedurisque = QLineEdit()
        self.DepotarisquedeVoletnaturedurisque.setPlaceholderText("Dépôt à risque de Vol et nature du risque")
        self.labelTauxdeBraquagepour100000habitants = QLabel("Taux de Braquage pour 100 000 habitants")
        self.TauxdeBraquagepour100000habitants = QLineEdit()
        self.TauxdeBraquagepour100000habitants.setPlaceholderText("Taux de Braquage pour 100 000 habitants")
        self.labelBraquages = QLabel("Braquages")
        self.Braquages = QLineEdit()
        self.Braquages.setPlaceholderText("Braquages")
        self.labelBraquagecible = QLabel("Braquage cible")
        self.Braquagecible = QLineEdit()
        self.Braquagecible.setPlaceholderText("Braquage cible")
        self.labelTauxdecambriolagepour100000habitants = QLabel("Taux de cambriolage pour 100 000 habitants")
        self.Tauxdecambriolagepour100000habitants = QLineEdit()
        self.Tauxdecambriolagepour100000habitants.setPlaceholderText("Taux de cambriolage pour 100 000 habitants")
        self.labelCambriolages = QLabel("Cambriolages")
        self.Cambriolages = QLineEdit()
        self.Cambriolages.setPlaceholderText("Cambriolages")
        self.labelDontCambriolageducoffre = QLabel("Dont Cambriolage du coffre")
        self.DontCambriolageducoffre = QLineEdit()
        self.DontCambriolageducoffre.setPlaceholderText("Dont Cambriolage du coffre")
        self.labelTotalEvenements = QLabel("Total Evènements")
        self.TotalEvenements = QLineEdit()
        self.TotalEvenements.setPlaceholderText("Total Evènements")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelTauxdevolspour100000habitants, 1, 0)
        self.layout.addWidget(self.Tauxdevolspour100000habitants, 1, 1, 1, 2)
        self.layout.addWidget(self.labelDepotarisquedeVoletnaturedurisque, 2, 0)
        self.layout.addWidget(self.DepotarisquedeVoletnaturedurisque, 2, 1, 1, 2)
        self.layout.addWidget(self.labelTauxdeBraquagepour100000habitants, 3, 0)
        self.layout.addWidget(self.TauxdeBraquagepour100000habitants, 3, 1, 1, 2)
        self.layout.addWidget(self.labelBraquages, 4, 0)
        self.layout.addWidget(self.Braquages, 4, 1, 1, 2)
        self.layout.addWidget(self.labelBraquagecible, 5, 0)
        self.layout.addWidget(self.Braquagecible, 5, 1, 1, 2)
        self.layout.addWidget(self.labelTauxdecambriolagepour100000habitants, 6, 0)
        self.layout.addWidget(self.Tauxdecambriolagepour100000habitants, 6, 1, 1, 2)
        self.layout.addWidget(self.labelCambriolages, 7, 0)
        self.layout.addWidget(self.Cambriolages, 7, 1, 1, 2)
        self.layout.addWidget(self.labelDontCambriolageducoffre, 8, 0)
        self.layout.addWidget(self.DontCambriolageducoffre, 8, 1, 1, 2)
        self.layout.addWidget(self.labelTotalEvenements, 9, 0)
        self.layout.addWidget(self.TotalEvenements, 9, 1, 1, 2)

        self.setLayout(self.layout)

##DONE17
class securite(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sécurité")

        self.setFixedSize(520, 400)

        # Create a scroll area widget
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFixedSize(500, 380)

        self.widget=QWidget()
        self.layout = QGridLayout()
        self.titre = QLabel("Sécurité")

        self.labelTypedeCoffre = QLabel("Type de Coffre")
        self.TypedeCoffre = QLineEdit()
        self.TypedeCoffre.setPlaceholderText("Type de Coffre")
        self.labelAutomateGLORY = QLabel("Automate GLORY")
        self.AutomateGLORY = QLineEdit()
        self.AutomateGLORY.setPlaceholderText("Automate GLORY")
        self.labelModeledeCoffre = QLabel("Modèle de Coffre")
        self.ModeledeCoffre = QLineEdit()
        self.ModeledeCoffre.setPlaceholderText("Modèle de Coffre")
        self.labelNombredePointsdeRamassageLOOMIS = QLabel("Nombre de Points de Ramassage LOOMIS")
        self.NombredePointsdeRamassageLOOMIS = QLineEdit()
        self.NombredePointsdeRamassageLOOMIS.setPlaceholderText("Nombre de Points de Ramassage LOOMIS")
        self.labelPneumatiqueCoffre = QLabel("Pneumatique Coffre")
        self.PneumatiqueCoffre = QLineEdit()
        self.PneumatiqueCoffre.setPlaceholderText("Pneumatique Coffre")
        self.labelAntennesAntiVol = QLabel("Antennes Anti-Vol")
        self.AntennesAntiVol = QLineEdit()
        self.AntennesAntiVol.setPlaceholderText("Antennes Anti-Vol")
        self.labelVideoCompleteenDepot = QLabel("Vidéo Complète en Dépôt")
        self.VideoCompleteenDepot = QLineEdit()
        self.VideoCompleteenDepot.setPlaceholderText("Vidéo Complète en Dépôt")
        self.labelVideodansleslocauxsociaux = QLabel("Vidéo dans les locaux sociaux")
        self.Videodansleslocauxsociaux = QLineEdit()
        self.Videodansleslocauxsociaux.setPlaceholderText("Vidéo dans les locaux sociaux")
        self.labelVideoenCaisses = QLabel("Vidéo en Caisses")
        self.VideoenCaisses = QLineEdit()
        self.VideoenCaisses.setPlaceholderText("Vidéo en Caisses")
        self.labelVideoenLogistique = QLabel("Vidéo en Logistique")
        self.VideoenLogistique = QLineEdit()
        self.VideoenLogistique.setPlaceholderText("Vidéo en Logistique")
        self.labelVideoauBati = QLabel("Vidéo au Bâti")
        self.VideoauBati = QLineEdit()
        self.VideoauBati.setPlaceholderText("Vidéo au Bâti")
        self.labelSocietedeSecurite2019 = QLabel("Société de Sécurité 2019")
        self.SocietedeSecurite2019 = QLineEdit()
        self.SocietedeSecurite2019.setPlaceholderText("Société de Sécurité 2019")
        self.labelCouthoraire2019 = QLabel("Coût horaire 2019")
        self.Couthoraire2019 = QLineEdit()
        self.Couthoraire2019.setPlaceholderText("Coût horaire 2019")
        self.labelNombredheuresGardiennage2017 = QLabel("Nombre d'heures Gardiennage 2017")
        self.NombredheuresGardiennage2017 = QLineEdit()
        self.NombredheuresGardiennage2017.setPlaceholderText("Nombre d'heures Gardiennage 2017")
        self.labelSocietedinterventiondeNuit = QLabel("Société d'intervention de Nuit")
        self.SocietedinterventiondeNuit = QLineEdit()
        self.SocietedinterventiondeNuit.setPlaceholderText("Société d'intervention de Nuit")
        self.labelDepotSprinkle = QLabel("Dépôt Sprinklé")
        self.DepotSprinkle = QLineEdit()
        self.DepotSprinkle.setPlaceholderText("Dépôt Sprinklé")
        self.labelSystemedeSecuriteIncendie = QLabel("Système de Sécurité Incendie")
        self.SystemedeSecuriteIncendie = QLineEdit()
        self.SystemedeSecuriteIncendie.setPlaceholderText("Système de Sécurité Incendie")
        self.labelModeledeCentraledAlarme = QLabel("Modèle de Centrale d'Alarme")
        self.ModeledeCentraledAlarme = QLineEdit()
        self.ModeledeCentraledAlarme.setPlaceholderText("Modèle de Centrale d'Alarme")
        self.labelInstallateurAlarme1 = QLabel("Installateur Alarme 1")
        self.InstallateurAlarme1 = QLineEdit()
        self.InstallateurAlarme1.setPlaceholderText("Installateur Alarme 1")
        self.labelInstallateurAlarme2 = QLabel("Installateur Alarme 2")
        self.InstallateurAlarme2 = QLineEdit()
        self.InstallateurAlarme2.setPlaceholderText("Installateur Alarme 2")
        self.labelTelesurveilleur = QLabel("Télésurveilleur")
        self.Telesurveilleur = QLineEdit()
        self.Telesurveilleur.setPlaceholderText("Télésurveilleur")
        self.labelTelesurveilleurTelephone = QLabel("Télésurveilleur Téléphone")
        self.TelesurveilleurTelephone = QLineEdit()
        self.TelesurveilleurTelephone.setPlaceholderText("Télésurveilleur Téléphone")
        self.labelGTB = QLabel("GTB")
        self.GTB = QLineEdit()
        self.GTB.setPlaceholderText("GTB")
        self.labelPortedaccesauxlocauxSociaux = QLabel("Porte d'accès aux locaux Sociaux")
        self.PortedaccesauxlocauxSociaux = QLineEdit()
        self.PortedaccesauxlocauxSociaux.setPlaceholderText("Porte d'accès aux locaux Sociaux")
        self.labelControledaccesparbadge = QLabel("Contrôle d'accès par badge")
        self.Controledaccesparbadge = QLineEdit()
        self.Controledaccesparbadge.setPlaceholderText("Contrôle d'accès par badge")
        self.labelControledaccesMigresurserveur = QLabel("Contrôle d'accès Migré sur serveur")
        self.ControledaccesMigresurserveur = QLineEdit()
        self.ControledaccesMigresurserveur.setPlaceholderText("Contrôle d'accès Migré sur serveur")
        self.labelDATIenMenuiserie = QLabel("DATI en Menuiserie")
        self.DATIenMenuiserie = QLineEdit()
        self.DATIenMenuiserie.setPlaceholderText("DATI en Menuiserie")
        self.labelGenerateurdefumeedansleLocalSecurite = QLabel("Générateur de fumée dans le Local Sécurité")
        self.GenerateurdefumeedansleLocalSecurite = QLineEdit()
        self.GenerateurdefumeedansleLocalSecurite.setPlaceholderText("Générateur de fumée dans le Local Sécurité")
        self.labelRalentisseurssurParking = QLabel("Ralentisseurs sur Parking")
        self.RalentisseurssurParking = QLineEdit()
        self.RalentisseurssurParking.setPlaceholderText("Ralentisseurs sur Parking")
        self.labelPlotsantibeliersurParking = QLabel("Plots anti bélier sur Parking")
        self.PlotsantibeliersurParking = QLineEdit()
        self.PlotsantibeliersurParking.setPlaceholderText("Plots anti bélier sur Parking")
        self.labelBarrieresdefermetureduParking = QLabel("Barrières de fermeture du Parking")
        self.BarrieresdefermetureduParking = QLineEdit()
        self.BarrieresdefermetureduParking.setPlaceholderText("Barrières de fermeture du Parking")
        self.labelBavoletsLogistique = QLabel("Bavolets Logistique")
        self.BavoletsLogistique = QLineEdit()
        self.BavoletsLogistique.setPlaceholderText("Bavolets Logistique")
        self.labelEclairageLED = QLabel("Eclairage LED")
        self.EclairageLED = QLineEdit()
        self.EclairageLED.setPlaceholderText("Eclairage LED")
        self.labelAscenseurs = QLabel("Ascenseurs")
        self.Ascenseurs = QLineEdit()
        self.Ascenseurs.setPlaceholderText("Ascenseurs")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelTypedeCoffre, 1, 0)
        self.layout.addWidget(self.TypedeCoffre, 1, 1, 1, 2)
        self.layout.addWidget(self.labelAutomateGLORY, 2, 0)
        self.layout.addWidget(self.AutomateGLORY, 2, 1, 1, 2)
        self.layout.addWidget(self.labelModeledeCoffre, 3, 0)
        self.layout.addWidget(self.ModeledeCoffre, 3, 1, 1, 2)
        self.layout.addWidget(self.labelNombredePointsdeRamassageLOOMIS, 4, 0)
        self.layout.addWidget(self.NombredePointsdeRamassageLOOMIS, 4, 1, 1, 2)
        self.layout.addWidget(self.labelPneumatiqueCoffre, 5, 0)
        self.layout.addWidget(self.PneumatiqueCoffre, 5, 1, 1, 2)
        self.layout.addWidget(self.labelAntennesAntiVol, 6, 0)
        self.layout.addWidget(self.AntennesAntiVol, 6, 1, 1, 2)
        self.layout.addWidget(self.labelVideoCompleteenDepot, 7, 0)
        self.layout.addWidget(self.VideoCompleteenDepot, 7, 1, 1, 2)
        self.layout.addWidget(self.labelVideodansleslocauxsociaux, 8, 0)
        self.layout.addWidget(self.Videodansleslocauxsociaux, 8, 1, 1, 2)
        self.layout.addWidget(self.labelVideoenCaisses, 9, 0)
        self.layout.addWidget(self.VideoenCaisses, 9, 1, 1, 2)
        self.layout.addWidget(self.labelVideoenLogistique, 10, 0)
        self.layout.addWidget(self.VideoenLogistique, 10, 1, 1, 2)
        self.layout.addWidget(self.labelVideoauBati, 11, 0)
        self.layout.addWidget(self.VideoauBati, 11, 1, 1, 2)
        self.layout.addWidget(self.labelSocietedeSecurite2019, 12, 0)
        self.layout.addWidget(self.SocietedeSecurite2019, 12, 1, 1, 2)
        self.layout.addWidget(self.labelCouthoraire2019, 13, 0)
        self.layout.addWidget(self.Couthoraire2019, 13, 1, 1, 2)
        self.layout.addWidget(self.labelNombredheuresGardiennage2017, 14, 0)
        self.layout.addWidget(self.NombredheuresGardiennage2017, 14, 1, 1, 2)
        self.layout.addWidget(self.labelSocietedinterventiondeNuit, 15, 0)
        self.layout.addWidget(self.SocietedinterventiondeNuit, 15, 1, 1, 2)
        self.layout.addWidget(self.labelDepotSprinkle, 16, 0)
        self.layout.addWidget(self.DepotSprinkle, 16, 1, 1, 2)
        self.layout.addWidget(self.labelSystemedeSecuriteIncendie, 17, 0)
        self.layout.addWidget(self.SystemedeSecuriteIncendie, 17, 1, 1, 2)
        self.layout.addWidget(self.labelModeledeCentraledAlarme, 18, 0)
        self.layout.addWidget(self.ModeledeCentraledAlarme, 18, 1, 1, 2)
        self.layout.addWidget(self.labelInstallateurAlarme1, 19, 0)
        self.layout.addWidget(self.InstallateurAlarme1, 19, 1, 1, 2)
        self.layout.addWidget(self.labelInstallateurAlarme2, 20, 0)
        self.layout.addWidget(self.InstallateurAlarme2, 20, 1, 1, 2)
        self.layout.addWidget(self.labelTelesurveilleur, 21, 0)
        self.layout.addWidget(self.Telesurveilleur, 21, 1, 1, 2)
        self.layout.addWidget(self.labelTelesurveilleurTelephone, 22, 0)
        self.layout.addWidget(self.TelesurveilleurTelephone, 22, 1, 1, 2)
        self.layout.addWidget(self.labelGTB, 23, 0)
        self.layout.addWidget(self.GTB, 23, 1, 1, 2)
        self.layout.addWidget(self.labelPortedaccesauxlocauxSociaux, 24, 0)
        self.layout.addWidget(self.PortedaccesauxlocauxSociaux, 24, 1, 1, 2)
        self.layout.addWidget(self.labelControledaccesparbadge, 25, 0)
        self.layout.addWidget(self.Controledaccesparbadge, 25, 1, 1, 2)
        self.layout.addWidget(self.labelControledaccesMigresurserveur, 26, 0)
        self.layout.addWidget(self.ControledaccesMigresurserveur, 26, 1, 1, 2)
        self.layout.addWidget(self.labelDATIenMenuiserie, 27, 0)
        self.layout.addWidget(self.DATIenMenuiserie, 27, 1, 1, 2)
        self.layout.addWidget(self.labelGenerateurdefumeedansleLocalSecurite, 28, 0)
        self.layout.addWidget(self.GenerateurdefumeedansleLocalSecurite, 28, 1, 1, 2)
        self.layout.addWidget(self.labelRalentisseurssurParking, 29, 0)
        self.layout.addWidget(self.RalentisseurssurParking, 29, 1, 1, 2)
        self.layout.addWidget(self.labelPlotsantibeliersurParking, 30, 0)
        self.layout.addWidget(self.PlotsantibeliersurParking, 30, 1, 1, 2)
        self.layout.addWidget(self.labelBarrieresdefermetureduParking, 31, 0)
        self.layout.addWidget(self.BarrieresdefermetureduParking, 31, 1, 1, 2)
        self.layout.addWidget(self.labelBavoletsLogistique, 32, 0)
        self.layout.addWidget(self.BavoletsLogistique, 32, 1, 1, 2)
        self.layout.addWidget(self.labelEclairageLED, 33, 0)
        self.layout.addWidget(self.EclairageLED, 33, 1, 1, 2)
        self.layout.addWidget(self.labelAscenseurs, 34, 0)
        self.layout.addWidget(self.Ascenseurs, 34, 1, 1, 2)

        self.widget.setLayout(self.layout)
        self.scroll_area.setWidget(self.widget)

##DONE18
class conceptCommercial(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Concept commercial")

        self.setFixedSize(520, 400)

        # Create a scroll area widget
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFixedSize(500, 380)
        
        self.widget = QWidget()


        self.layout = QGridLayout()
        self.titre = QLabel("Concept commercial")

        self.labelFacadeROUGE = QLabel("Façade ROUGE")
        self.FacadeROUGE = QLineEdit()
        self.FacadeROUGE.setPlaceholderText("Façade ROUGE")
        self.labelRevitalisationRemodling = QLabel("Revitalisation Remodling")
        self.RevitalisationRemodling = QLineEdit()
        self.RevitalisationRemodling.setPlaceholderText("Revitalisation Remodling")
        self.labelDRIVE = QLabel("DRIVE")
        self.DRIVE = QLineEdit()
        self.DRIVE.setPlaceholderText("DRIVE")
        self.labelDRIVEdate = QLabel("DRIVE date")
        self.DRIVEdate = QLineEdit()
        self.DRIVEdate.setPlaceholderText("DRIVE date")
        self.labelIDepot = QLabel("I Dépôt")
        self.IDepot = QLineEdit()
        self.IDepot.setPlaceholderText("I Dépôt")
        self.labelLivraisonLocaleInStore = QLabel("Livraison Locale In Store")
        self.LivraisonLocaleInStore = QLineEdit()
        self.LivraisonLocaleInStore.setPlaceholderText("Livraison Locale In Store")
        self.labelLivraisonLocaleInStoreDate = QLabel("Livraison Locale In Store Date")
        self.LivraisonLocaleInStoreDate = QLineEdit()
        self.LivraisonLocaleInStoreDate.setPlaceholderText("Livraison Locale In Store Date")
        self.labelZRM = QLabel("ZRM")
        self.ZRM = QLineEdit()
        self.ZRM.setPlaceholderText("ZRM")
        self.labelZRMDate = QLabel("ZRM Date")
        self.ZRMDate = QLineEdit()
        self.ZRMDate.setPlaceholderText("ZRM Date")
        self.labelContenuZRM = QLabel("Contenu ZRM")
        self.ContenuZRM = QLineEdit()
        self.ContenuZRM.setPlaceholderText("Contenu ZRM")
        self.labelimprimantededieeZRM = QLabel("imprimante dédiée ZRM")
        self.imprimantededieeZRM = QLineEdit()
        self.imprimantededieeZRM.setPlaceholderText("imprimante dédiée ZRM")
        self.labelEASIEROULEGACY = QLabel("EASIER OU LEGACY")
        self.EASIEROULEGACY = QLineEdit()
        self.EASIEROULEGACY.setPlaceholderText("EASIER OU LEGACY")
        self.labelBacasable = QLabel("Bac à sable")
        self.Bacasable = QLineEdit()
        self.Bacasable.setPlaceholderText("Bac à sable")
        self.labelFournisseurSable = QLabel("Fournisseur Sable")
        self.FournisseurSable = QLineEdit()
        self.FournisseurSable.setPlaceholderText("Fournisseur Sable")
        self.labelTremie = QLabel("Trémie")
        self.Tremie = QLineEdit()
        self.Tremie.setPlaceholderText("Trémie")
        self.labelGodethydraulique = QLabel("Godet hydraulique")
        self.Godethydraulique = QLineEdit()
        self.Godethydraulique.setPlaceholderText("Godet hydraulique")
        self.labelGodetmecanique = QLabel("Godet mécanique")
        self.Godetmecanique = QLineEdit()
        self.Godetmecanique.setPlaceholderText("Godet mécanique")
        self.labelBetonalatoupie = QLabel("Béton à la toupie")
        self.Betonalatoupie = QLineEdit()
        self.Betonalatoupie.setPlaceholderText("Béton à la toupie")
        self.labelchargeuses = QLabel("chargeuses")
        self.chargeuses = QLineEdit()
        self.chargeuses.setPlaceholderText("chargeuses")
        self.labeldatederetourmails = QLabel("date de retour mails")
        self.datederetourmails = QLineEdit()
        self.datederetourmails.setPlaceholderText("date de retour mails")
        self.labelIDEPOT = QLabel("I DEPOT")
        self.IDEPOT = QLineEdit()
        self.IDEPOT.setPlaceholderText("I DEPOT")
        self.labelDecoupeBois = QLabel("Découpe Bois")
        self.DecoupeBois = QLineEdit()
        self.DecoupeBois.setPlaceholderText("Découpe Bois")
        self.labelTestRenfortequipe = QLabel("Test Renfort équipe")
        self.TestRenfortequipe = QLineEdit()
        self.TestRenfortequipe.setPlaceholderText("Test Renfort équipe")
        self.labelTestsurMesureplacards = QLabel("Test sur Mesure placards")
        self.TestsurMesureplacards = QLineEdit()
        self.TestsurMesureplacards.setPlaceholderText("Test sur Mesure placards")
        self.labelTestsurMesureMenuiserie = QLabel("Test sur Mesure Menuiserie")
        self.TestsurMesureMenuiserie = QLineEdit()
        self.TestsurMesureMenuiserie.setPlaceholderText("Test sur Mesure Menuiserie")
        self.labelTestTelephoniesurIP = QLabel("Test Téléphonie sur IP")
        self.TestTelephoniesurIP = QLineEdit()
        self.TestTelephoniesurIP.setPlaceholderText("Test Téléphonie sur IP")
        self.labelTestRETENCYTrackingClients = QLabel("Test RETENCY Tracking Clients")
        self.TestRETENCYTrackingClients = QLineEdit()
        self.TestRETENCYTrackingClients.setPlaceholderText("Test RETENCY Tracking Clients")
        self.labelPresentoiraDalles = QLabel("Présentoir à Dalles")
        self.PresentoiraDalles = QLineEdit()
        self.PresentoiraDalles.setPlaceholderText("Présentoir à Dalles")
        self.labelShowroomSalledeBains = QLabel("Showroom Salle de Bains")
        self.ShowroomSalledeBains = QLineEdit()
        self.ShowroomSalledeBains.setPlaceholderText("Showroom Salle de Bains")
        self.labelShowroomSalledeBainsDate = QLabel("Showroom Salle de Bains Date")
        self.ShowroomSalledeBainsDate = QLineEdit()
        self.ShowroomSalledeBainsDate.setPlaceholderText("Showroom Salle de Bains Date")
        self.labelErgosquelette = QLabel("Ergosquelette")
        self.Ergosquelette = QLineEdit()
        self.Ergosquelette.setPlaceholderText("Ergosquelette")
        self.labelTranspaletteciseaux = QLabel("Transpalette ciseaux")
        self.Transpaletteciseaux = QLineEdit()
        self.Transpaletteciseaux.setPlaceholderText("Transpalette ciseaux")
        self.labelPlaquesteflonees = QLabel("Plaques téflonées")
        self.Plaquesteflonees = QLineEdit()
        self.Plaquesteflonees.setPlaceholderText("Plaques téflonées")
        self.labelCagesPalettes = QLabel("Cages Palettes")
        self.CagesPalettes = QLineEdit()
        self.CagesPalettes.setPlaceholderText("Cages Palettes")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelFacadeROUGE, 1, 0)
        self.layout.addWidget(self.FacadeROUGE, 1, 1, 1, 2)
        self.layout.addWidget(self.labelRevitalisationRemodling, 2, 0)
        self.layout.addWidget(self.RevitalisationRemodling, 2, 1, 1, 2)
        self.layout.addWidget(self.labelDRIVE, 3, 0)
        self.layout.addWidget(self.DRIVE, 3, 1, 1, 2)
        self.layout.addWidget(self.labelDRIVEdate, 4, 0)
        self.layout.addWidget(self.DRIVEdate, 4, 1, 1, 2)
        self.layout.addWidget(self.labelIDepot, 5, 0)
        self.layout.addWidget(self.IDepot, 5, 1, 1, 2)
        self.layout.addWidget(self.labelLivraisonLocaleInStore, 6, 0)
        self.layout.addWidget(self.LivraisonLocaleInStore, 6, 1, 1, 2)
        self.layout.addWidget(self.labelLivraisonLocaleInStoreDate, 7, 0)
        self.layout.addWidget(self.LivraisonLocaleInStoreDate, 7, 1, 1, 2)
        self.layout.addWidget(self.labelZRM, 8, 0)
        self.layout.addWidget(self.ZRM, 8, 1, 1, 2)
        self.layout.addWidget(self.labelZRMDate, 9, 0)
        self.layout.addWidget(self.ZRMDate, 9, 1, 1, 2)
        self.layout.addWidget(self.labelContenuZRM, 10, 0)
        self.layout.addWidget(self.ContenuZRM, 10, 1, 1, 2)
        self.layout.addWidget(self.labelimprimantededieeZRM, 11, 0)
        self.layout.addWidget(self.imprimantededieeZRM, 11, 1, 1, 2)
        self.layout.addWidget(self.labelEASIEROULEGACY, 12, 0)
        self.layout.addWidget(self.EASIEROULEGACY, 12, 1, 1, 2)
        self.layout.addWidget(self.labelBacasable, 13, 0)
        self.layout.addWidget(self.Bacasable, 13, 1, 1, 2)
        self.layout.addWidget(self.labelFournisseurSable, 14, 0)
        self.layout.addWidget(self.FournisseurSable, 14, 1, 1, 2)
        self.layout.addWidget(self.labelTremie, 15, 0)
        self.layout.addWidget(self.Tremie, 15, 1, 1, 2)
        self.layout.addWidget(self.labelGodethydraulique, 16, 0)
        self.layout.addWidget(self.Godethydraulique, 16, 1, 1, 2)
        self.layout.addWidget(self.labelGodetmecanique, 17, 0)
        self.layout.addWidget(self.Godetmecanique, 17, 1, 1, 2)
        self.layout.addWidget(self.labelBetonalatoupie, 18, 0)
        self.layout.addWidget(self.Betonalatoupie, 18, 1, 1, 2)
        self.layout.addWidget(self.labelchargeuses, 19, 0)
        self.layout.addWidget(self.chargeuses, 19, 1, 1, 2)
        self.layout.addWidget(self.labeldatederetourmails, 20, 0)
        self.layout.addWidget(self.datederetourmails, 20, 1, 1, 2)
        self.layout.addWidget(self.labelIDEPOT, 21, 0)
        self.layout.addWidget(self.IDEPOT, 21, 1, 1, 2)
        self.layout.addWidget(self.labelDecoupeBois, 22, 0)
        self.layout.addWidget(self.DecoupeBois, 22, 1, 1, 2)
        self.layout.addWidget(self.labelTestRenfortequipe, 23, 0)
        self.layout.addWidget(self.TestRenfortequipe, 23, 1, 1, 2)
        self.layout.addWidget(self.labelTestsurMesureplacards, 24, 0)
        self.layout.addWidget(self.TestsurMesureplacards, 24, 1, 1, 2)
        self.layout.addWidget(self.labelTestsurMesureMenuiserie, 25, 0)
        self.layout.addWidget(self.TestsurMesureMenuiserie, 25, 1, 1, 2)
        self.layout.addWidget(self.labelTestTelephoniesurIP, 26, 0)
        self.layout.addWidget(self.TestTelephoniesurIP, 26, 1, 1, 2)
        self.layout.addWidget(self.labelTestRETENCYTrackingClients, 27, 0)
        self.layout.addWidget(self.TestRETENCYTrackingClients, 27, 1, 1, 2)
        self.layout.addWidget(self.labelPresentoiraDalles, 28, 0)
        self.layout.addWidget(self.PresentoiraDalles, 28, 1, 1, 2)
        self.layout.addWidget(self.labelShowroomSalledeBains, 29, 0)
        self.layout.addWidget(self.ShowroomSalledeBains, 29, 1, 1, 2)
        self.layout.addWidget(self.labelShowroomSalledeBainsDate, 30, 0)
        self.layout.addWidget(self.ShowroomSalledeBainsDate, 30, 1, 1, 2)
        self.layout.addWidget(self.labelErgosquelette, 31, 0)
        self.layout.addWidget(self.Ergosquelette, 31, 1, 1, 2)
        self.layout.addWidget(self.labelTranspaletteciseaux, 32, 0)
        self.layout.addWidget(self.Transpaletteciseaux, 32, 1, 1, 2)
        self.layout.addWidget(self.labelPlaquesteflonees, 33, 0)
        self.layout.addWidget(self.Plaquesteflonees, 33, 1, 1, 2)
        self.layout.addWidget(self.labelCagesPalettes, 34, 0)
        self.layout.addWidget(self.CagesPalettes, 34, 1, 1, 2)

        self.widget.setLayout(self.layout)
        self.scroll_area.setWidget(self.widget)
    
##DONE : Attention labelContexteComptage-TCUENTO19
class divers(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Divers")
        self.layout = QGridLayout()
        self.titre = QLabel("Divers")

        self.labelLesclientspenetrentdanslazonedesbureauxpouraccederauxtoilettes = QLabel("Les clients pénétrent dans la zone des bureaux pour accéder aux toilettes")
        self.Lesclientspenetrentdanslazonedesbureauxpouraccederauxtoilettes = QLineEdit()
        self.Lesclientspenetrentdanslazonedesbureauxpouraccederauxtoilettes.setPlaceholderText("Les clients pénétrent dans la zone des bureaux pour accéder aux toilettes")
        self.labelSocietedeNettoyage = QLabel("Société de Nettoyage")
        self.SocietedeNettoyage = QLineEdit()
        self.SocietedeNettoyage.setPlaceholderText("Société de Nettoyage")
        self.labelFenetresaletagedesbureaux = QLabel("Fenêtres à l'étage des bureaux")
        self.Fenetresaletagedesbureaux = QLineEdit()
        self.Fenetresaletagedesbureaux.setPlaceholderText("Fenêtres à l'étage des bureaux")
        self.labelComptageClients = QLabel("Comptage Clients")
        self.ComptageClients = QLineEdit()
        self.ComptageClients.setPlaceholderText("Comptage Clients")
        self.labelContexteComptageTCUENTO = QLabel("Contexte Comptage - T CUENTO")
        self.ContexteComptageTCUENTO = QLineEdit()
        self.ContexteComptageTCUENTO.setPlaceholderText("Contexte Comptage - T CUENTO")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelLesclientspenetrentdanslazonedesbureauxpouraccederauxtoilettes, 1, 0)
        self.layout.addWidget(self.Lesclientspenetrentdanslazonedesbureauxpouraccederauxtoilettes, 1, 1, 1, 2)
        self.layout.addWidget(self.labelSocietedeNettoyage, 2, 0)
        self.layout.addWidget(self.SocietedeNettoyage, 2, 1, 1, 2)
        self.layout.addWidget(self.labelFenetresaletagedesbureaux, 3, 0)
        self.layout.addWidget(self.Fenetresaletagedesbureaux, 3, 1, 1, 2)
        self.layout.addWidget(self.labelComptageClients, 4, 0)
        self.layout.addWidget(self.ComptageClients, 4, 1, 1, 2)
        self.layout.addWidget(self.labelContexteComptageTCUENTO, 5, 0)
        self.layout.addWidget(self.ContexteComptageTCUENTO, 5, 1, 1, 2)

        self.setLayout(self.layout)

##DONE : ° remplacé par deg20
class numCommercant(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Numéro commercant")
        self.layout = QGridLayout()
        self.titre = QLabel("Numéro commercant")

        self.labelNdegTVA = QLabel("N° TVA")
        self.NdegTVA = QLineEdit()
        self.NdegTVA.setPlaceholderText("N° TVA")
        self.labelNdegCommercant = QLabel("N° Commerçant")
        self.NdegCommercant = QLineEdit()
        self.NdegCommercant.setPlaceholderText("N° Commerçant")
        self.labelCodeAbonneTransax = QLabel("Code Abonné Transax")
        self.CodeAbonneTransax = QLineEdit()
        self.CodeAbonneTransax.setPlaceholderText("Code Abonné Transax")
        self.labelCodeVendeurCACF = QLabel("Code Vendeur CACF")
        self.CodeVendeurCACF = QLineEdit()
        self.CodeVendeurCACF.setPlaceholderText("Code Vendeur CACF")
        self.labelCodeIPCACF = QLabel("Code IP CACF")
        self.CodeIPCACF = QLineEdit()
        self.CodeIPCACF.setPlaceholderText("Code IP CACF")
        self.labelNdegCommercantdrive = QLabel("N° Commerçant drive")
        self.NdegCommercantdrive = QLineEdit()
        self.NdegCommercantdrive.setPlaceholderText("N° Commerçant drive")
        self.labelNdegCommercantPayementsanscontact = QLabel("N° Commerçant Payement sans contact")
        self.NdegCommercantPayementsanscontact = QLineEdit()
        self.NdegCommercantPayementsanscontact.setPlaceholderText("N° Commerçant Payement sans contact")
        self.labelCodeIBANBNP = QLabel("Code IBAN/ BNP")
        self.CodeIBANBNP = QLineEdit()
        self.CodeIBANBNP.setPlaceholderText("Code IBAN/ BNP")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelNdegTVA, 1, 0)
        self.layout.addWidget(self.NdegTVA, 1, 1, 1, 2)
        self.layout.addWidget(self.labelNdegCommercant, 2, 0)
        self.layout.addWidget(self.NdegCommercant, 2, 1, 1, 2)
        self.layout.addWidget(self.labelCodeAbonneTransax, 3, 0)
        self.layout.addWidget(self.CodeAbonneTransax, 3, 1, 1, 2)
        self.layout.addWidget(self.labelCodeVendeurCACF, 4, 0)
        self.layout.addWidget(self.CodeVendeurCACF, 4, 1, 1, 2)
        self.layout.addWidget(self.labelCodeIPCACF, 5, 0)
        self.layout.addWidget(self.CodeIPCACF, 5, 1, 1, 2)
        self.layout.addWidget(self.labelNdegCommercantdrive, 6, 0)
        self.layout.addWidget(self.NdegCommercantdrive, 6, 1, 1, 2)
        self.layout.addWidget(self.labelNdegCommercantPayementsanscontact, 7, 0)
        self.layout.addWidget(self.NdegCommercantPayementsanscontact, 7, 1, 1, 2)
        self.layout.addWidget(self.labelCodeIBANBNP, 8, 0)
        self.layout.addWidget(self.CodeIBANBNP, 8, 1, 1, 2)

        self.setLayout(self.layout)

##DONE21
class colissimo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Colissimo")
        self.layout = QGridLayout()
        self.titre = QLabel("Colissimo")

        self.labelCompteIdentifiant = QLabel("Compte Identifiant")
        self.CompteIdentifiant = QLineEdit()
        self.CompteIdentifiant.setPlaceholderText("Compte Identifiant")
        self.labelMotdepasse = QLabel("Mot de passe")
        self.Motdepasse = QLineEdit()
        self.Motdepasse.setPlaceholderText("Mot de passe")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelCompteIdentifiant, 1, 0)
        self.layout.addWidget(self.CompteIdentifiant, 1, 1, 1, 2)
        self.layout.addWidget(self.labelMotdepasse, 2, 0)
        self.layout.addWidget(self.Motdepasse, 2, 1, 1, 2)

        self.setLayout(self.layout)

##DONE22
class accidentTravail(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Accident travail")
        self.layout = QGridLayout()
        self.titre = QLabel("Accident travail")

        self.labelTauxdAT2015 = QLabel("Taux d'AT 2015")
        self.TauxdAT2015 = QLineEdit()
        self.TauxdAT2015.setPlaceholderText("Taux d'AT 2015")
        self.labelTauxdAT2016 = QLabel("Taux d'AT 2016")
        self.TauxdAT2016 = QLineEdit()
        self.TauxdAT2016.setPlaceholderText("Taux d'AT 2016")
        self.labelTauxdAT2017 = QLabel("Taux d'AT 2017")
        self.TauxdAT2017 = QLineEdit()
        self.TauxdAT2017.setPlaceholderText("Taux d'AT 2017")
        self.labelTauxdAT2018 = QLabel("Taux d'AT 2018")
        self.TauxdAT2018 = QLineEdit()
        self.TauxdAT2018.setPlaceholderText("Taux d'AT 2018")
        self.labelTauxdAT2019 = QLabel("Taux d'AT 2019")
        self.TauxdAT2019 = QLineEdit()
        self.TauxdAT2019.setPlaceholderText("Taux d'AT 2019")
        self.labelTauxdAT2020 = QLabel("Taux d'AT 2020")
        self.TauxdAT2020 = QLineEdit()
        self.TauxdAT2020.setPlaceholderText("Taux d'AT 2020")
        self.labelTauxdAT2021 = QLabel("Taux d'AT 2021")
        self.TauxdAT2021 = QLineEdit()
        self.TauxdAT2021.setPlaceholderText("Taux d'AT 2021")

        self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)

        self.layout.addWidget(self.labelTauxdAT2015, 1, 0)
        self.layout.addWidget(self.TauxdAT2015, 1, 1, 1, 2)
        self.layout.addWidget(self.labelTauxdAT2016, 2, 0)
        self.layout.addWidget(self.TauxdAT2016, 2, 1, 1, 2)
        self.layout.addWidget(self.labelTauxdAT2017, 3, 0)
        self.layout.addWidget(self.TauxdAT2017, 3, 1, 1, 2)
        self.layout.addWidget(self.labelTauxdAT2018, 4, 0)
        self.layout.addWidget(self.TauxdAT2018, 4, 1, 1, 2)
        self.layout.addWidget(self.labelTauxdAT2019, 5, 0)
        self.layout.addWidget(self.TauxdAT2019, 5, 1, 1, 2)
        self.layout.addWidget(self.labelTauxdAT2020, 6, 0)
        self.layout.addWidget(self.TauxdAT2020, 6, 1, 1, 2)
        self.layout.addWidget(self.labelTauxdAT2021, 7, 0)
        self.layout.addWidget(self.TauxdAT2021, 7, 1, 1, 2)

        self.setLayout(self.layout)

##TODO: Voir l'efficacité
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
        self.listeCodeBRICO = self.sheet["Code BRICO"].values.tolist()

        for i in range (len(self.listedesdepots)):
            self.listedesdepots[i] = str(self.listeCodeBRICO[i]) + "-" + self.listedesdepots[i]

        self.listedepots.addItem('')
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
        self.listedepots2.currentIndexChanged.connect(self.chargementTableau)
        self.affichageDep = QTableView()

        self.boutonConfirmation = QPushButton("Confirmer la modification")
        self.boutonConfirmation.clicked.connect(self.choix)

        self.layout2.addWidget(self.titre2, 0, 1)
        self.layout2.addWidget(self.listedepots2, 1, 1,)
        self.layout2.addWidget(self.affichageDep, 2, 0, 1, 3)
        self.layout2.addWidget(self.boutonConfirmation, 3, 1)


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
        self.setFixedSize(720,440)           

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
    
##TODO: Supprimer la ligne dans le dataframe
class suppWindow(QWidget):
    sheet = pd.DataFrame
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Suppression d'un dépôt")
        self.title = QLabel("Suppression d'un dépôt")

        self.setFixedSize(200,80)

        self.stack = QStackedWidget()
        self.i= 5
        ##TODO: Widget 1 -> Sélection du dépot
        self.widget1 = QWidget()
        self.layout1 = QVBoxLayout()

        self.titre = QLabel("Sélection du dépôt à supprimer")
        self.listedepots = QComboBox()
        self.listedesdepots = self.sheet["Dépôt"].values.tolist()
        self.listeCodeBRICO = self.sheet["Code BRICO"].values.tolist()

        for i in range (len(self.listedesdepots)):
            self.listedesdepots[i] = str(self.listeCodeBRICO[i]) + "-" + self.listedesdepots[i]

        self.listedepots.addItem('')
        self.listedepots.addItems(self.listedesdepots)
        self.listedepots.currentIndexChanged.connect(self.selectionDepot)

        self.layout1.addWidget(self.titre)
        self.layout1.addWidget(self.listedepots)

        self.widget1.setLayout(self.layout1)
        
        ##TODO: Widget 2 -> Affichage du dépot + possibilité de toujours le modifier
        self.widget2 = QWidget()
        self.layout2 = QGridLayout()
        self.titre2 = QLabel("Sélection du dépôt à supprimer")
        self.listedepots2 = QComboBox()
        self.listedepots2.addItems(self.listedesdepots)
        self.listedepots2.currentIndexChanged.connect(self.chargementTableau)
        self.affichageDep = QTableView()

        self.boutonConfirmation = QPushButton("Supprimer le dépôt")
        self.boutonConfirmation.clicked.connect(self.suppression)
        #, Qt.AlignmentFlag.AlignCenter
        self.layout2.addWidget(self.titre2, 0, 1)
        self.layout2.addWidget(self.listedepots2, 1, 1,)
        self.layout2.addWidget(self.affichageDep, 2, 0, 1, 3)
        self.layout2.addWidget(self.boutonConfirmation, 3, 1)

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

##TODO: Transmission du sheet_tri 
class demandeChangement(QWidget):
    sheet = pd.DataFrame
    def __init__(self):
        super().__init__()
        self.setFixedSize(720,440)

        self.setWindowTitle("Requête changement donnée")

        self.layout = QGridLayout()

        self.Titre = QLabel("Demande de changement de données")

        self.texte = QLabel("Modifier les données dans le tableau ci-dessous")
        self.table = QTableView()
        self.model = pandasEditableModel(self.sheet)
        self.table.setModel(self.model)
        self.table.resizeColumnsToContents()

        self.valider = QPushButton("Valider la demande de changement de données")
        self.valider.clicked.connect(self.transmettre)
        
        self.layout.addWidget(self.Titre, 0, 1)
        self.layout.addWidget(self.texte, 1, 1)
        self.layout.addWidget(self.table, 2, 0, 1, 3)
        self.layout.addWidget(self.valider, 3, 1)

        self.setLayout(self.layout)

    def transmettre(self):
        ##Envoyer le set de données quelque part avec le user qui demande et la date
        
        self.close()

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
    
    form = logWindow()
    main = mainWindow()
    form.show()

    sys.exit(app.exec())
