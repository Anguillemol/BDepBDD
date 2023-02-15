#Program to write into Excel

import pandas as pd
import numpy as np
import PySimpleGUI as sg

username = ""
password = ""

####################################################### Récupération des données depuis la BDD  ####################################################### 

#Commande locale
sheet = pd.read_excel('C:/Users/lucas/Test.xlsx', sheet_name='Données')

#Commande Sharepoint
#
#
#

####################################################### TRAITEMENT DES DONNEES  #######################################################

#Pour supprimer des colonnes
sheet_drop = sheet.drop("Lettres", axis = 1)

#Pour ajouter des colonnes
sheet_add = sheet.copy(deep=True)
sheet_add["Ajout1"] = ""
sheet_add["Ajout2"] = np.nan

#Filtre pour récupérer uniquement les valeurs qui ont un id impair
sheet_filter = sheet[sheet.Id%2 != 0]

#Affichage
print("Sheet")
print(sheet)
print("Sheet drop")
print(sheet_drop)
print("Sheet add")
print(sheet_add)
print("Sheet filter")
print(sheet_filter)

########################################################## SAISIE DES DONNEES ########################################################
with pd.ExcelWriter('C:/Users/lucas/Test.xlsx', mode='a', if_sheet_exists='replace') as writer:
    sheet_filter.to_excel(writer, sheet_name='Filtre')
    sheet_drop.to_excel(writer, sheet_name='Drop')
    sheet_add.to_excel(writer, sheet_name='Add')

########################################################### AFFICHAGE TEMPORAIRE ######################################################

"""

headings = list(sheet.columns)
print(headings)
values = sheet.values.tolist()

sg.theme("DarkBlue3")
sg.set_options(font=("Courier New", 16))

layout = [[sg.Table(values = values, headings = headings,
    # Set column widths for empty record of table
    auto_size_columns=False,
    col_widths=list(map(lambda x:len(x)+1, headings)))]]

window = sg.Window('Sample excel file',  layout)
event, value = window.read()


#https://statisticsglobe.com/dataframe-manipulation-using-pandas-python
"""
