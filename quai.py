import pandas as pd
from tabulate import tabulate

file_path = "C:/Users/lenovo/Desktop/semaine d'optimisation/Données GWD 2026 VF.xlsx"

try:
    df1 = pd.read_excel(file_path, sheet_name='Données Transporteurs', nrows=6, usecols="A:D")

    df2 = pd.read_excel(file_path, sheet_name='Données Transporteurs', skiprows=8, nrows=13, usecols="A:D")

    print("Data loaded successfully:")
    print(df1.head())
    print(df2.head())

except FileNotFoundError:
    print(f"Error: The file '{file_path}' was not found.")
except Exception as e:
    print(f"An error occurred: {e}")





# Affiche la première table avec des bordures et les en-têtes
print(tabulate(df1, headers='keys', tablefmt='psql'))
print(tabulate(df2, headers='keys', tablefmt='psql'))