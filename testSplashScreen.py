import sys
import time
import io
import pandas as pd
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QProgressBar, QLabel, QFrame, QHBoxLayout, QVBoxLayout
from PyQt6.QtCore import Qt, QTimer
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File 

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

        # center labels
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
        self.labelLoading.setText('loading...')

        self.labelSubDescription = QLabel(self.frame)
        self.labelSubDescription.resize(self.width() - 10, 50)
        self.labelSubDescription.move(0, self.labelDescription.y()+ 50)
        self.labelSubDescription.setObjectName('LabelSubDesc')
        self.labelSubDescription.setText('')
        self.labelSubDescription.setAlignment(Qt.AlignmentFlag.AlignCenter)

    def loading(self):
        self.progressBar.setValue(self.counter)

        if self.counter == int(self.n * 0.3):
            self.labelDescription.setText('<strong>Working on Task #2</strong>')
        elif self.counter == int(self.n * 0.6):
            self.labelDescription.setText('<strong>Working on Task #3</strong>')
        elif self.counter >= self.n:
            self.timer.stop()
            self.close()

            time.sleep(1)

            self.myApp = MyApp()
            self.myApp.show()

        self.counter += 1



class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.window_width, self.window_height = 1200, 800
        self.setMinimumSize(self.window_width, self.window_height)

        layout = QVBoxLayout()
        self.setLayout(layout)


if __name__ == '__main__':
    # don't auto scale when drag app to a different monitor.
    # QApplication.setAttribute(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    
    app = QApplication(sys.argv)
    """app.setStyleSheet('''
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
    """
    splash = SplashScreen()
    splash.show()

    app.processEvents()
    splash.labelDescription.setText('<strong>Connexion à Sharepoint</strong>')

    #####ETAPE 1#####
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

    splash.progressBar.setValue(10)
    splash.labelDescription.setText('<strong>Récupération fichier 1</strong>')

    #####ETAPE 2#####
    responseBDD = File.open_binary(ctx, bdd_URL)
    print("Reponse trouvée")
    bytes_file_obj_bdd = io.BytesIO()
    bytes_file_obj_bdd.write(responseBDD.content)
    bytes_file_obj_bdd.seek(0)
    print ("BDD chargée")

    splash.progressBar.setValue(20)
    splash.labelDescription.setText('<strong>Récupération fichier 2</strong>')
    
    #####ETAPE 3#####
    responseMDP = File.open_binary(ctx, mdp_URL)
    print("Reponse trouvée")
    bytes_file_obj_mdp = io.BytesIO()
    bytes_file_obj_mdp.write(responseMDP.content)
    bytes_file_obj_mdp.seek(0)
    print ("MDP chargés")

    splash.progressBar.setValue(30)
    splash.labelDescription.setText('<strong>Récupération fichier 3</strong>')

    #####ETAPE 4#####
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

    #####ETAPE 5#####
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
    try:
        sys.exit(app.exec())
    except SystemExit:
        print('Closing Window...')