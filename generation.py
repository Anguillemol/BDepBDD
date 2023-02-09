from unidecode import unidecode

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

