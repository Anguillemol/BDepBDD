import sys
import io
import tempfile
import os
import shutil
import openpyxl as wb
import pandas as pd
import numpy as np

from pathlib import Path
from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File 
from datetime import datetime
from  cryptography.fernet import Fernet

with open('key.key', 'rb') as key_file:
    key = key_file.read()

with open('config.cfg', 'rb') as config_file:
    encrypted_user = config_file.readline()
    encrypted_password = config_file.readline()
fernet = Fernet(key)
username = fernet.decrypt(encrypted_user).decode()
password = fernet.decrypt(encrypted_password).decode()

test_team_site_url = "https://sgzkl.sharepoint.com/sites/BricoDepot"
bdd_URL = "/sites/BricoDepot/Shared%20Documents/Donnees/BDD.xlsx"
mdp_URL = "/sites/BricoDepot/Shared%20Documents/Donnees/MDP.xlsx"
requete_URL = "/sites/BricoDepot/Shared%20Documents/Donnees/REQ.xlsx"

ctx = ClientContext(test_team_site_url).with_credentials(ClientCredential(username, password))
web = ctx.web
ctx.load(web).execute_query()
print ("Connexion à Sharepoint réussie")

responseBDD = File.open_binary(ctx, bdd_URL)
print("Reponse trouvée")
bytes_file_obj_bdd = io.BytesIO()
bytes_file_obj_bdd.write(responseBDD.content)
bytes_file_obj_bdd.seek(0)
print("BDD chargée")

bdd = bytes_file_obj_bdd

xlsx_file = pd.ExcelFile(bdd)
sheet_lst = xlsx_file.sheet_names

sheet_Globale = pd.DataFrame()
##TODO: !!!!!
###PAS BIN ATTENTION POOUAITNTA NTAIONTOIANT FAIRE UNE CONCAT sur axis 1
column_set = set()
bool_premier = True
bool_region2022 = True
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



print(sheet_Globale)
sheet_Globale.to_excel("outTest.xlsx")