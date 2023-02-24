import pandas as pd

# Créer le premier DataFrame
df1 = pd.DataFrame({'col1': [1, 2, 3], 'col2': [4, 5, 6], 'col3': [7, 8, 9]})

# Copier les colonnes de df1 dans df2
df2 = pd.DataFrame()
df2['col1'] = df1.iloc[:, 0]
df2['col2'] = df1.iloc[:, 1]

# Afficher les DataFrames pour vérifier les résultats
print(df1)
print(df2)
