import pandas as pd
import openpyxl
from openpyxl import load_workbook


#Chargement fichier Excel
filename = 'test/bddTest.xlsx'
wb = load_workbook(filename)

#Extraction onglet requete
for sheet_name in wb.sheetnames:
    print(sheet_name)
    if sheet_name != 'Req':
        sheet = wb[sheet_name]
        colonnesees = [sheet.cell(row=1, column=c).value for c in range(1, sheet.max_column + 1)]
        print(colonnesees)
        
        colonne_prime = None
        print(list(sheet.values))
        iColonne = 0
        for colonne in sheet.iter_cols(min_col=1, values_only=True):
            iColonne +=1
            if colonne[0] == "Prime":
                colonne_prime = list(colonne)
                break

        if colonne_prime is not None:
            for i in range(1, len(colonne_prime)):
                colonne_prime[i] += 1
            
            for i, valeur in enumerate(colonne_prime):
                sheet.cell(row=i+1, column=iColonne, value=valeur)
        print("nouvelle list")
        print(list(sheet.values))
#Chargement des autres onglets dans les pandas


#Modif dans les dataframes pandas


#Ecriture dans les onglets
wb.save(filename)