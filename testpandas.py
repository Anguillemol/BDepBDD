import pandas as pd

# Créer les dataframes
df1 = pd.DataFrame({'A': [1, 2, 3], 'B': [4, 5, 6], 'C': [7, 8, 9]})
df2 = pd.DataFrame({'A': [10, 11, 12], 'B': [13, 14, 15], 'C': [16, 17, 18], 'utilisateur': ['User1', 'User2', 'User3'], 'date demande': ['2022-01-01', '2022-01-02', '2022-01-03']})

# Concaténer les dataframes
df_concat = pd.concat([df1, df2], axis=0)

# Afficher le dataframe concaténé
print(df_concat)