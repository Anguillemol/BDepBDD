from unidecode import unidecode
import pandas as pd
import secrets
import string

letters = string.ascii_letters
digits = string.digits


alphabet = letters + digits

pwd_length = 8

"""
def traitement(input):
    trait = unidecode(input)
    trait = trait.replace(' ','')
    trait = trait.replace("'", "")
    trait = trait.replace("(","")
    trait = trait.replace(")","")
    trait = trait.replace("/", "")
    trait = trait.replace("-", "")
    return trait

ligne1 = "Taux d'AT 2015	Taux d'AT 2016	Taux d'AT 2017	Taux d'AT 2018	Taux d'AT 2019	Taux d'AT 2020	Taux d'AT 2021"
ligne1 = ligne1.replace("	", ",")

listedata= ligne1.split(",")
print (ligne1)
print("")

for i in range(len(listedata)):
    print('self.label'+traitement(listedata[i])+' = QLabel("' + listedata[i] + '")')
    print('self.'+traitement(listedata[i])+ ' = QLineEdit()')
    print('self.'+traitement(listedata[i])+'.setPlaceholderText("' + listedata[i] + '")')

print("")
print("self.layout.addWidget(self.titre,0,0,1,3,Qt.AlignmentFlag.AlignCenter)")
print("")
for i in range(len(listedata)):
    num = str(i+1)
    print('self.layout.addWidget(self.label'+traitement(listedata[i])+', '+num+', 0)')
    print('self.layout.addWidget(self.'+traitement(listedata[i])+', '+num+', 1, 1, 2)')

print("")
print("self.setLayout(self.layout)")
"""


##Generateur username et password

dfRegion = pd.DataFrame(columns=['username','password'])
dfDepot = pd.DataFrame(columns=['username','password'])
dfAdmin = pd.DataFrame(columns=['username','password'])

def traitementUsername(prenomNom):
    premiereLettre = prenomNom[0]
    separationNom = prenomNom.find(" ")
    cinqLettresNom = prenomNom[separationNom + 1:separationNom + 6]
    username = premiereLettre+cinqLettresNom
    return username


def creaMDP():
    password = ""
    for i in range(pwd_length):
        password += ''.join(secrets.choice(alphabet))
    return password

stringRegion = "Jean-Marie PELUT,Stéphane MURAIS,Cyril ROBINET,Pascal FREVOL,Thierry GIANELLA,Guillaume FARVACQUE,Patrick PAPOT,Olivier FRUCHART,Pierre LAVILLAT,Christophe ROYER,Pierre Olivier LARUE,Wylli TINTIN"
listeRegion = stringRegion.split(",")

stringDepot = "MICHEL CHIPY,YANNICK PIERRE,SYLVAIN CANONNE,ERIC JOULAUD,URBINO ESTEVES,ERIC PERROT,MANUELLA CASTRO,XAVIER MOREAU,PHILIPPE GARLASCHI,PHILIPPE SALVATORE,HERVE FRIBOULET,DAVID CATRY,LAURENCE VACCARIZI,ROMAIN BALLOY / ERIC PETRA,BASTIEN FORTE,FRANCIS BLANCAFORT,DAVID REBUFFE,STEPHANE VAREILLE,FREDERIC THEVENIAUT,EMMANUEL ROQUES,LAETITIA PEIXOTO,ROMAIN BALLOY,JEAN PIERRE HOAREAU,KARIM LOUAHAB,OLIVER DAVID,BRUNO CAINAUD,JOACHIM GUIHENEUF,NICOLAS BALLOIR,CHRISTOPHE FERDINANDE,DAVID DE SOUSA,ANTONY OBIOLS,ANTHONY BEAUBAT,RACHID ANNAB,JEAN FRANCOIS ROITEL,STEPHANE MILLIET,ALLAL NAKIB,OLIVIER HORNOY,SEBASTIEN CASADESUS,SONIA PORCU,ANNE  FOUREL,PAUL MARCONI,FRANCK RAFFIN,FABRICE MERSCH,MATHIEU LIENHARDT,ANTHONY ELOY,MARC BULTEL,FREDERIC MERRE LEVEQUE,CLAUDE UNTEREINER,JEAN-PAUL TEETEN,FRANCK DELANGE,DJAMEL OULDSLIMANE,SEBASTIEN GUIBERT,AHMED NESSATI,FLAVIEN ARCHAMBAULT,YVAN MAHIEU,SEBASTIEN ZIMMERMANN,FABRICE BONIEC,FRANCK LEMAIRE,LAURENT BRUNET,VANESSA QUENSIERE,KEVIN VIVIER,PASCAL DEMARECAUX,STEPHANE CAMUS,WILLIAM EDMOND,JEAN-ROBERT DRUART,LUDOVIC VANCUTSEM,PASCAL LEDOUBLE,OLIVIER MEILLIEZ,SEBASTIEN DALICHOUX,DAVID CHARLE,NICOLAS YON,GUILLAUME BRUNET,FABRICE MARQUES,ALAIN PIERRE,JOSE OLIVEIRA,LAURENCE MODESTE,MATHIEU HAULTCOEUR,FRANCOIS DELAUNAY,YOHANN MOURTOUX,XAVIER LEGRAND,REMY VANDENBERGHE,YOUCEF ELMECHTA,NATHALIE CUENOT,RACHID BENYKHLEF,GILLES LEJEAN,LAURENT FIRMIN,BRUNO MERLAND,LUDOVIC KOLTALO,BOUBEKEUR BAKOUR,DOMINIQUE VETTIER,EMMANUEL BROSSAY,FREDERIC PLESSE,ERWAN GOURIOU,NAJIB BOUCHNAK,ERIC LONEGRO,FREDERIC LACROIX,CEDRIC PUPIER,GUILLAUME CHESNAIS,CHRISTOPHE DALLEMAGNE,MATHIEU VERRIER,JEAN MARC ANSOTTE,LAURENT TUDAL,SEBASTIEN QUENOT,JULIEN MUTIN,MICHEL VERLAINE,HAMID ASSIOUI,PASCAL THELLIER,CEDRIC PIAZZA,BERTRAND BIGNAN,OLIVIER BELET,BERTRAND COTTEAU,MICHEL LORIA,PAUL WITKAMP,FRANTZ DECIEUX,LOUIS TRACOL,DAMIEN STEFANIAK,CHRISTOPHE ORTU,FABRICE MARTIN,ALEXIS ROTTIER,THIERRY COUASNON,PATRICK HERRERO"
listeDepot = stringDepot.split(",")

stringAdmin = ""
listeAdmin = stringAdmin.split(",")

for i in range(len(listeRegion)):
    username = traitementUsername(listeRegion[i])
    password = creaMDP()
    new_df = pd.DataFrame([[username, password]], columns=['username', 'password'])
    dfRegion = pd.concat([dfRegion, new_df], axis = 0, ignore_index=True)


for i in range(len(listeDepot)):
    username = traitementUsername(listeDepot[i])
    password = creaMDP()
    new_df = pd.DataFrame([[username, password]], columns=['username', 'password'])
    dfDepot = pd.concat([dfDepot, new_df], axis = 0, ignore_index=True)

if stringAdmin != "":
    for i in range(len(listeAdmin)):
        username = traitementUsername(listeAdmin[i])
        password = creaMDP()
        new_df = pd.DataFrame([[username, password]], columns=['username', 'password'])
        dfAdmin = pd.concat([dfAdmin, new_df], axis = 0, ignore_index=True)

print(dfDepot)
print(dfRegion)

dfDepot.to_csv("dfDepot.csv", encoding='utf-8', index = False)
dfRegion.to_csv("dfRegion.csv", encoding='utf-8', index=False)