from PyQt6.QtWidgets import QApplication, QMainWindow, QLabel, QProgressBar, QVBoxLayout, QWidget
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import sys, io, shutil, tempfile, os, datetime, time
import pandas as pd
from unidecode import unidecode
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File 


class WorkerThread(QThread):
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)

    def run(self):
        # Connexion à SharePoint
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
        self.update_progress(10)

        # Récupération des fichiers
        responseBDD = File.open_binary(ctx, bdd_URL)
        print("Reponse trouvée")
        bytes_file_obj_bdd = io.BytesIO()
        bytes_file_obj_bdd.write(responseBDD.content)
        bytes_file_obj_bdd.seek(0)
        print("BDD chargée")
        self.update_progress(23)

        responseMDP = File.open_binary(ctx, mdp_URL)
        print("Reponse trouvée")
        bytes_file_obj_mdp = io.BytesIO()
        bytes_file_obj_mdp.write(responseMDP.content)
        bytes_file_obj_mdp.seek(0)
        print ("MDP chargés")
        self.update_progress(36)

        responseREQ = File.open_binary(ctx, requete_URL)
        print("Reponse trouvée")
        bytes_file_obj_req = io.BytesIO()
        bytes_file_obj_req.write(responseREQ.content)
        bytes_file_obj_req.seek(0)
        print ("Requêtes chargées")
        self.update_progress(50)

        ##### Storing the data streams #####
        bdd = bytes_file_obj_bdd
        mdp = bytes_file_obj_mdp
        req = bytes_file_obj_req
        
        # Création de la sheet_globale
        print("Lecture stream BDD pour création du set de données")
        xlsx_file = pd.ExcelFile(bdd)
        sheet_lst = xlsx_file.sheet_names

        sheet_Globale = pd.DataFrame()
        ##TODO: vérif
        bool_premier = True
        for i in range(len(sheet_lst)):
            progression = int(round(50 + (i/(len(sheet_lst))*50)))
            print(progression)
            if sheet_lst[i] != 'Accueil' and sheet_lst[i] != 'BDD':
                print (sheet_lst[i])
                workSheet = pd.read_excel(bdd, sheet_name=sheet_lst[i])
                common_col = ["Code BRICO", "Code EASIER", "Dépôt", "Région 2022"]
                if bool_premier == True:
                    sheet_Globale = workSheet
                    bool_premier=False
                else:
                    sheet_Globale = pd.merge(sheet_Globale, workSheet, on=common_col)
            self.update_progress(progression)
        self.update_progress(100)
        # Émission d'un signal pour indiquer que le travail est terminé
        self.finished_signal.emit()

    def update_progress(self, value):
        self.progress_signal.emit(value)

class SplashScreen(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Splash Screen")
        self.setFixedSize(300, 200)

        layout = QVBoxLayout()
        self.label = QLabel("Chargement des données...", alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        layout.addWidget(self.progress_bar)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.worker_thread = WorkerThread()
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.show_main_window)
        self.worker_thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def show_main_window(self):
        self.main_window = MainWindow()
        self.main_window.show()
        self.close()

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Main Window")
        self.setFixedSize(400, 300)

if __name__ == "__main__":
    app = QApplication([])
    splash = SplashScreen()
    splash.show()
    app.exec()

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