from flask import Flask, jsonify
import subprocess
import sys
import os


app = Flask(__name__)

processes = {}  # Dictionnaire pour stocker les processus en cours

# Route pour Logs
@app.route('/logs', methods=['GET'])
def get_logs():
    try:
        with open('central_rpa.log', 'r') as log_file:
            logs = log_file.read()
        return logs, 200  # Retourne les logs en tant que texte brut
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    
# Routes pour Affranchigo
@app.route('/start-affranchigo', methods=['POST'])
def start_affranchigo():
    if 'Affranchigo' in processes:
        return jsonify({"status": "RPA Affranchigo déjà en cours"}), 400
    # Récupérer le chemin absolu du répertoire parent (GUI)
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # Chemin vers le dossier GUI
    main_py_path = os.path.join(base_dir, 'data', 'main.py')  # Chemin absolu vers main.py dans GUI/data/

    try:
        # Lancer main.py en subprocess avec Python
        process = subprocess.Popen([sys.executable, main_py_path, 'Affranchigo'])
        
        # Stocker le processus en cours
        processes['Affranchigo'] = process
        return jsonify({"status": "RPA Extraction démarré"}), 200
    except FileNotFoundError as e:
        return jsonify({"status": f"Erreur : Fichier non trouvé - {e}"}), 500
    except Exception as e:
        return jsonify({"status": f"Erreur : {e}"}), 500

@app.route('/stop-affranchigo', methods=['POST'])
def stop_affranchigo():
    if 'Affranchigo' in processes:
        processes['Affranchigo'].terminate()
        processes['Affranchigo'].wait()
        del processes['Affranchigo']
        return jsonify({"status": "Affranchigo RPA arrêté"}), 200
    return jsonify({"status": "Aucun RPA Affranchigo en cours"}), 400

@app.route('/status-affranchigo', methods=['GET'])
def status_affranchigo():
    if 'Affranchigo' in processes and processes['Affranchigo'].poll() is None:
        return jsonify({"status": "Affranchigo RPA en cours"}), 200
    return jsonify({"status": "Affranchigo RPA non démarré"}), 200

# Routes pour Dématerialisation
@app.route('/start-dematerialisation', methods=['POST'])
def start_dematerialisation():
    if 'CasDematerialisation' in processes:
        return jsonify({"status": "RPA CasDematerialisation déjà en cours"}), 400
    # Récupérer le chemin absolu du répertoire parent (GUI)
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # Chemin vers le dossier GUI
    main_py_path = os.path.join(base_dir, 'data', 'main.py')  # Chemin absolu vers main.py dans GUI/data/

    try:
        # Lancer main.py en subprocess avec Python
        process = subprocess.Popen([sys.executable, main_py_path, 'CasDematerialisation'])
        
        # Stocker le processus en cours
        processes['CasDematerialisation'] = process
        return jsonify({"status": "RPA Extraction démarré"}), 200
    except FileNotFoundError as e:
        return jsonify({"status": f"Erreur : Fichier non trouvé - {e}"}), 500
    except Exception as e:
        return jsonify({"status": f"Erreur : {e}"}), 500

@app.route('/stop-dematerialisation', methods=['POST'])
def stop_dematerialisation():
    if 'dematerialisation' in processes:
        processes['dematerialisation'].terminate()
        processes['dematerialisation'].wait()
        del processes['dematerialisation']
        return jsonify({"status": "RPA Dématerialisation arrêté"}), 200
    return jsonify({"status": "Aucun RPA Dématerialisation en cours"}), 400

@app.route('/status-dematerialisation', methods=['GET'])
def status_dematerialisation():
    if 'dematerialisation' in processes and processes['dematerialisation'].poll() is None:
        return jsonify({"status": "RPA Dématerialisation en cours"}), 200
    return jsonify({"status": "RPA Dématerialisation non démarré"}), 200

# Routes pour Extraction Odyssée
@app.route('/start-extraction', methods=['POST'])
def start_extraction():
    if 'extraction' in processes:
        return jsonify({"status": "RPA Extraction déjà en cours"}), 400
    
    # Récupérer le chemin absolu du répertoire parent (GUI)
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # Chemin vers le dossier GUI
    main_py_path = os.path.join(base_dir, 'data', 'main.py')  # Chemin absolu vers main.py dans GUI/data/

    try:
        # Lancer main.py en subprocess avec Python
        process = subprocess.Popen([sys.executable, main_py_path, 'Extraction'])
        
        # Stocker le processus en cours
        processes['extraction'] = process
        return jsonify({"status": "RPA Extraction démarré"}), 200
    except FileNotFoundError as e:
        return jsonify({"status": f"Erreur : Fichier non trouvé - {e}"}), 500
    except Exception as e:
        return jsonify({"status": f"Erreur : {e}"}), 500

@app.route('/stop-extraction', methods=['POST'])
def stop_extraction():
    if 'extraction' in processes:
        processes['extraction'].terminate()
        processes['extraction'].wait()
        del processes['extraction']
        return jsonify({"status": "RPA Extraction arrêté"}), 200
    return jsonify({"status": "Aucun RPA Extraction en cours"}), 400

@app.route('/status-extraction', methods=['GET'])
def status_extraction():
    if 'extraction' in processes and processes['extraction'].poll() is None:
        return jsonify({"status": "RPA Extraction en cours"}), 200
    return jsonify({"status": "RPA Extraction non démarré"}), 200

# Routes pour Portage de Repas
@app.route('/start-portage-repas', methods=['POST'])
def start_portage_repas():
    if 'portage_repas' in processes:
        return jsonify({"status": "RPA Portage de Repas déjà en cours"}), 400
    process = subprocess.Popen([sys.executable, 'data/main.py', 'PortageRepas'])
    processes['portage_repas'] = process
    return jsonify({"status": "RPA Portage de Repas démarré"}), 200

@app.route('/stop-portage-repas', methods=['POST'])
def stop_portage_repas():
    if 'portage_repas' in processes:
        processes['portage_repas'].terminate()
        processes['portage_repas'].wait()
        del processes['portage_repas']
        return jsonify({"status": "RPA Portage de Repas arrêté"}), 200
    return jsonify({"status": "Aucun RPA Portage de Repas en cours"}), 400

@app.route('/status-portage-repas', methods=['GET'])
def status_portage_repas():
    if 'portage_repas' in processes and processes['portage_repas'].poll() is None:
        return jsonify({"status": "RPA Portage de Repas en cours"}), 200
    return jsonify({"status": "RPA Portage de Repas non démarré"}), 200

# Routes pour Seres
@app.route('/start-seres', methods=['POST'])
def start_seres():
    if 'seres' in processes:
        return jsonify({"status": "Seres RPA déjà en cours"}), 400
    process = subprocess.Popen([sys.executable, 'data/main.py', 'Seres'])
    processes['seres'] = process
    return jsonify({"status": "Seres RPA démarré"}), 200

@app.route('/stop-seres', methods=['POST'])
def stop_seres():
    if 'seres' in processes:
        processes['seres'].terminate()
        processes['seres'].wait()
        del processes['seres']
        return jsonify({"status": "Seres RPA arrêté"}), 200
    return jsonify({"status": "Aucun RPA Seres en cours"}), 400

@app.route('/status-seres', methods=['GET'])
def status_seres():
    if 'seres' in processes and processes['seres'].poll() is None:
        return jsonify({"status": "Seres RPA en cours"}), 200
    return jsonify({"status": "Seres RPA non démarré"}), 200


if __name__ == "__main__":
    app.run(debug=True)
