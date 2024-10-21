import sqlite3
import os

# Fonction pour lire la table 'infos_extraites'
def read_db_table(db_path="C:/Users/PHSX095/OneDrive - LA POSTE GROUPE/Documents/GUI/data/data_traitement/data_extraction.db"):
    """
    Lit toutes les lignes de la table 'infos_extraites' dans la base de données SQLite.
    Retourne les résultats sous forme de liste de dictionnaires.
    """
    try:
        # Vérifier si la base de données existe
        if not os.path.exists(db_path):
            print(f"Base de données non trouvée à {db_path}")
            return

        # Se connecter à la base de données
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()

        # Exécuter la requête pour lire toutes les lignes de la table 'infos_extraites'
        cursor.execute("SELECT * FROM infos_extraites")
        rows = cursor.fetchall()

        # Récupérer les noms des colonnes
        column_names = [description[0] for description in cursor.description]

        # Convertir les résultats en une liste de dictionnaires
        results = []
        for row in rows:
            results.append(dict(zip(column_names, row)))

        conn.close()

        # Afficher les résultats
        print(f"Nombre de lignes lues : {len(results)}")
        for result in results:
            print(result)

    except sqlite3.Error as e:
        print(f"Erreur lors de la lecture de la base de données : {e}")

# Point d'entrée pour exécuter la fonction
if __name__ == "__main__":
    read_db_table()
