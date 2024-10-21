import os
import json
import pandas as pd

def extract_contrat_numbers_to_json(excel_path, json_path):
    # Charger le fichier Excel ou CSV
    df = pd.read_excel(excel_path)
    
    # Créer un dictionnaire avec un format simple
    contrat_numbers = {str(i): num for i, num in enumerate(df["Contrat Nb"])}
    
    # Sauvegarder en JSON
    with open(json_path, 'w') as json_file:
        json.dump(contrat_numbers, json_file)
