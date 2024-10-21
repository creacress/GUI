import pandas as pd
import sqlite3

def execute_sql_on_excel(excel_file, sql_query, sheet_name=0):
    """
    Exécute une requête SQL sur un fichier Excel en utilisant SQLite comme backend temporaire.

    :param excel_file: Chemin vers le fichier Excel.
    :param sql_query: Requête SQL à exécuter.
    :param sheet_name: Nom ou numéro de l'onglet dans l'Excel à charger (par défaut le premier).
    :return: Résultat de la requête SQL sous forme de DataFrame pandas.
    """
    try:
        # Charger le fichier Excel dans un DataFrame pandas
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # Connexion à une base de données SQLite en mémoire
        conn = sqlite3.connect(':memory:')
        
        # Sauvegarder les données dans une table SQLite
        df.to_sql('excel_data', conn, index=False, if_exists='replace')

        # Exécuter la requête SQL sur les données
        result = pd.read_sql_query(sql_query, conn)

        # Fermer la connexion à la base de données
        conn.close()

        return result

    except Exception as e:
        print(f"Erreur lors de l'exécution de la requête SQL sur le fichier Excel: {e}")
        return None
