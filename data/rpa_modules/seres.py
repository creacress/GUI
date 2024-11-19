import os
import time
import json
import pandas as pd
import threading
from dotenv import load_dotenv
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor, as_completed
from rpa_modules.debug import setup_logger
from selenium.webdriver.common.action_chains import ActionChains
from rpa_modules.data_processing import extract_contrat_numbers_to_json
from dash import Dash, dcc, html, dash_table
from dash.dependencies import Input, Output
import plotly.express as px

load_dotenv()
# Siret destinataire corrigé selon Prmedi
class SeresRPA:

    def __init__(self, pool, logger=None):
        self.pool = pool
        self.logger = logger or setup_logger("seres_case.log")
        self.url = "https://portail.e-facture.net/saml/saml-login.php?nomSP=ARTIMON_PROD"
        self.processed_count = 0
        self.success_count = 0
        self.failure_count = 0
        self.total_duration = 0
        self.results = []
        self.error_logs = []
        self.lock = threading.Lock()
    
    def update_metrics(self, success, duration, numero_facture, message):
        """
        Met à jour les métriques et les résultats pour le tableau de bord.
        """
        with self.lock:
            self.processed_count += 1
            if success:
                self.success_count += 1
            else:
                self.failure_count += 1
            self.total_duration += duration
            self.results.append({
                "Numéro de Contrat": numero_facture,
                "Statut": "Succès" if success else "Échec",
                "Message": message,
                "Durée (s)": duration
            })


    def log_error(self, numero_facture, message):
        """
        Log une erreur et l'ajoute à la liste des logs d'erreurs.
        """
        self.logger.error(message)
        self.error_logs.append({"NumeroFacture": numero_facture, "Error": message})
    
    def start_dashboard(self):
        """
        Initialise un tableau de bord Dash pour le suivi en temps réel.
        """
        app = Dash(__name__)

        app.layout = html.Div([
            html.H1("Tableau de bord SeresRPA"),
            dash_table.DataTable(
                id="result-table",
                columns=[
                    {"name": "Numéro de Contrat", "id": "Numéro de Contrat"},
                    {"name": "Statut", "id": "Statut"},
                    {"name": "Message", "id": "Message"},
                    {"name": "Durée (s)", "id": "Durée (s)"}
                ],
                data=[],
                style_table={"overflowX": "auto"},
                style_cell={"textAlign": "left"},
                style_header={"fontWeight": "bold"},
            ),
            dcc.Graph(id="success-failure-pie"),
            html.Div([
                html.H4("Nombre total de contrats traités :"),
                html.P(id="total-count"),
                html.H4("Temps moyen de traitement :"),
                html.P(id="average-duration"),
            ]),
            dcc.Interval(id="interval-component", interval=1000, n_intervals=0)
        ])

        @app.callback(
            [Output("result-table", "data"),
             Output("success-failure-pie", "figure"),
             Output("total-count", "children"),
             Output("average-duration", "children")],
            [Input("interval-component", "n_intervals")]
        )
        def update_dashboard(n):
            total = self.processed_count
            average_duration = self.total_duration / total if total > 0 else 0
            pie_fig = px.pie(
                names=["Succès", "Échecs"],
                values=[self.success_count, self.failure_count],
                title="Répartition des contrats traités"
            )
            with self.lock:
                return self.results, pie_fig, f"{total}", f"{average_duration:.2f} secondes"

        app.run_server(debug=True, use_reloader=False)

    
    def process_json_files(self, file_path):
        """
        Traite un fichier JSON contenant les numéros de contrat et renvoie une liste de ces numéros.
        """
        self.logger.debug("Traitement du fichier JSON pour les contrats...")
        numeros_contrat = []

        if os.path.exists(file_path):
            with open(file_path, "r") as file:
                data = json.load(file)
                numeros_contrat = list(data.values())

        return numeros_contrat


    def check_page_loaded(self, driver):
        try:
            WebDriverWait(driver, 30).until(
                EC.url_to_be(self.url)
            )
            self.logger.info(f"Page {self.url} chargée avec succès.")
        except TimeoutException:
            self.logger.error(f"La page {self.url} n'a pas été chargée correctement.")
            raise

    def click_login_button(self, driver):
        try:
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div.container > div > div.col-xs-12.col-md-6 > div > form > button"))
            )
            login_button = driver.find_element(By.CSS_SELECTOR, "body > div.container > div > div.col-xs-12.col-md-6 > div > form > button")
            login_button.click()
            self.logger.info("Le bouton de connexion a été cliqué avec succès.")
        except Exception as e:
            self.logger.error(f"Erreur lors du clic sur le bouton de connexion: {e}")

    def login(self, driver, wait, identifiant, mot_de_passe):
        self.logger.debug("Tentative de connexion...")
        try:
            # Étape : Vérification que la page est bien chargée
            self.check_page_loaded(driver)

            input_identifiant = wait.until(EC.presence_of_element_located((By.ID, "login")))
            input_identifiant.clear()
            input_identifiant.send_keys(identifiant)

            input_mot_de_passe = wait.until(EC.presence_of_element_located((By.ID, "acct_pass")))
            input_mot_de_passe.clear()
            input_mot_de_passe.send_keys(mot_de_passe)

            self.click_login_button(driver)
            self.logger.info("Connexion effectuée avec succès.")
        except TimeoutException:
            self.logger.error("Erreur : Les champs de connexion ne sont pas disponibles.")
        except Exception as e:
            self.logger.error(f"Problème lors de la connexion : {e}")
    
    def save_non_modifiable(driver, numero_facture, json_filename):
        try:
            # Créer le répertoire pour les captures d'écran si nécessaire
            screenshot_filename = f"screenshots/erreur_{numero_facture}.png"
            os.makedirs(os.path.dirname(screenshot_filename), exist_ok=True)
            # Charger les données JSON existantes ou créer un nouveau fichier si nécessaire
            if os.path.exists(json_filename):
                with open(json_filename, "r") as json_file:
                    data = json.load(json_file)
            else:
                data = []

            # Ajouter le numéro de facture à la liste
            data.append({"numero_facture": numero_facture})

            # Sauvegarder les nouvelles données dans le fichier JSON
            with open(json_filename, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print(f"Numéro de facture {numero_facture} sauvegardé dans le fichier {json_filename}")
        
        except Exception as e:
            print(f"Erreur lors de la sauvegarde du fichier ou de la capture d'écran : {e}")

    def is_error_page(self, driver):
        try:
            # Localiser l'élément qui pourrait indiquer la page d'erreur
            error_element = driver.find_element(By.CSS_SELECTOR, "#content > h2:nth-child(9)")
            # Vérifier si le texte correspond à "Envoyer le rapport d'erreur"
            if error_element.text == "Envoyer le rapport d'erreur":
                return True
            return False
        except Exception:
            return False
    
    def write_comment(self, driver, comment_text):
        try:
            # Localiser la zone de texte par son sélecteur CSS
            comment_box = driver.find_element(By.CSS_SELECTOR, "#m_commentaire_interne")
            # Effacer le texte existant (si nécessaire) et écrire le commentaire
            comment_box.clear()
            comment_box.send_keys(comment_text)
            self.logger.info("Commentaire écrit avec succès.")
        except Exception as e:
            self.logger.error(f"Erreur lors de l'écriture du commentaire : {e}")
    
    def click_validate_button(self, driver, numero_facture):
        try:
            # Localiser le bouton à l'aide du sélecteur CSS
            validate_button = driver.find_element(By.CSS_SELECTOR, "#validate")
            # Scroller jusqu'au bouton si nécessaire et cliquer
            ActionChains(driver).move_to_element(validate_button).click().perform()
            self.logger.info("Bouton de validation cliqué avec succès.")
        except Exception as e:
            self.logger.error(f"Erreur lors du clic sur le bouton de validation Valider: {e}")
            self.save_non_modifiable(numero_facture, "validation_button_erreur.json")
    
    def click_button_by_text(self, driver, button_text):
        """
        Clique sur un bouton contenant un texte spécifique.
        """
        try:
            self.logger.debug(f"Recherche du bouton avec le texte '{button_text}'...")
            
            # Localiser les boutons contenant le texte spécifié
            buttons = driver.find_elements(By.XPATH, f"//button[contains(text(), '{button_text}')]")

            if buttons:
                for button in buttons:
                    # Vérifier si le bouton est visible et interactif
                    if button.is_displayed() and button.is_enabled():
                        self.logger.info(f"Bouton '{button_text}' trouvé et cliqué.")
                        ActionChains(driver).move_to_element(button).click().perform()
                        return
                self.logger.warning(f"Aucun bouton visible avec le texte '{button_text}'.")
            else:
                self.logger.error(f"Aucun bouton trouvé avec le texte '{button_text}'.")
                raise Exception(f"Bouton '{button_text}' introuvable.")
        except Exception as e:
            self.logger.error(f"Erreur lors du clic sur le bouton '{button_text}' : {e}")
            raise

    def click_and_validate_modal(self, driver, main_button_text, modal_button_text):
        """
        Gère le clic sur un bouton principal et le clic sur un bouton dans une modale qui s'affiche ensuite.
        
        :param driver: Instance du WebDriver.
        :param main_button_text: Texte du bouton principal qui ouvre la modale.
        :param modal_button_text: Texte du bouton à cliquer dans la modale.
        """
        try:
            # Étape 1 : Clic sur le bouton principal
            self.logger.info(f"Tentative de clic sur le bouton principal '{main_button_text}'...")
            self.click_button_by_text(driver, main_button_text)

            # Étape 2 : Attendre l'apparition de la modale
            self.logger.info(f"Attente de l'apparition de la modale après clic sur '{main_button_text}'...")
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.XPATH, f"//div[contains(@class, 'modal') and contains(@style, 'display: block')]"))
            )

            # Étape 3 : Clic sur le bouton dans la modale
            self.logger.info(f"Tentative de clic sur le bouton '{modal_button_text}' dans la modale...")
            self.click_button_by_text(driver, modal_button_text)

            self.logger.info("Clic sur le bouton dans la modale réussi.")
        except TimeoutException:
            self.logger.error(f"Modale après le clic sur '{main_button_text}' n'est pas apparue.")
            raise
        except Exception as e:
            self.logger.error(f"Erreur lors de la gestion de la modale : {e}")
            raise

    def click_rejets_aife(self, driver):
        """
        Fonction pour cliquer sur la div "Rejets AIFE".
        :param driver: L'instance WebDriver en cours.
        :param logger: Le logger pour enregistrer les logs.
        """
        try:
            self.logger.info("Recherche de la div contenant 'Rejets AIFE'...")

            # Attendre que l'élément avec le texte "Rejets AIFE" soit visible
            wait = WebDriverWait(driver, 10)
            div_element = wait.until(EC.presence_of_element_located((By.XPATH, "//h1[text()='Rejets AIFE']/ancestor::div[@class='panel-body']")))

            # Cliquer sur la div englobante
            div_element.click()

            self.logger.info("Clic sur 'Rejets AIFE' effectué avec succès.")
        
        except TimeoutException:
            self.logger.error("Le bouton 'Rejets AIFE' n'a pas été trouvé sur la page.")
        except Exception as e:
            self.logger.error(f"Erreur lors du clic sur 'Rejets AIFE' : {e}")

    def enter_num_facture(self, driver, numero_facture):
        """
        Saisit le numéro de facture dans le champ de recherche, lance la recherche et nettoie le champ.
        """
        try:
            self.logger.info(f"Saisie du numéro facture : {numero_facture}")

            # Attendre que le champ soit interactif (cliquable)
            input_numfacture = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "gs_NUMFACTURE"))
            )

            # Effacer le champ et entrer le numéro de facture
            input_numfacture.clear()
            input_numfacture.send_keys(numero_facture)

            # Valider la recherche avec Entrée
            input_numfacture.send_keys(Keys.ENTER)
            self.logger.debug("Appui sur Entrée pour lancer la recherche.")

            # Alternative : Si ENTER échoue, essayer de cliquer sur le bouton de recherche
            try:
                search_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@type='submit']"))
                )
                search_button.click()
                self.logger.debug("Clic sur le bouton de recherche en complément de l'appui sur Entrée.")
            except TimeoutException:
                self.logger.warning("Le bouton de recherche n'a pas été trouvé, uniquement Entrée utilisé.")

            # Vérifier que les résultats sont chargés (table avec des lignes)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//table[contains(@id, 'list-documents')]//tr"))
            )

            self.logger.info(f"Numéro facture {numero_facture} saisi et recherche lancée avec succès.")

        except TimeoutException:
            self.logger.error("Le champ de saisie pour 'NUMFACTURE' n'a pas été trouvé ou n'est pas interactif.")
        except Exception as e:
            self.logger.error(f"Erreur lors de la saisie du numéro facture : {e}")
            if "element not interactable" in str(e):
                self.logger.warning("Nouvelle tentative après 'element not interactable'.")
                time.sleep(1)  # Attendre brièvement avant de relancer
                self.enter_num_facture(driver, numero_facture)
        finally:
            # Nettoyer le champ après le traitement
            try:
                input_numfacture.clear()
                self.logger.debug("Champ de saisie nettoyé après le traitement.")
            except Exception as e:
                self.logger.warning(f"Impossible de nettoyer le champ après le traitement : {e}")



    def select_row_by_facture(self, driver, numero_facture):
        """
        Sélectionne la ligne du tableau correspondant au numéro facture fourni et clique dessus.
        :param driver: L'instance WebDriver.
        :param numero_facture: Le numéro facture à rechercher dans le tableau.
        """
        try:
            self.logger.info(f"Sélection de la ligne avec le numéro facture : {numero_facture}...")

            # Attendre que le tableau contenant les documents soit visible
            wait = WebDriverWait(driver, 20)
            table = wait.until(EC.visibility_of_element_located((By.XPATH, "//table[contains(@id, 'list-documents')]")))
            time.sleep(2)  # Ajout d'un délai court pour s'assurer que les lignes sont chargées

            # Récupérer toutes les lignes du tableau (exclure la première ligne si elle est un header)
            rows = table.find_elements(By.XPATH, ".//tr[not(contains(@class, 'jqgfirstrow'))]")
            self.logger.debug(f"Nombre de lignes trouvées dans le tableau : {len(rows)}")

            # Comparaison pour vérifier que le numéro de facture est contenu dans une cellule
            for row in rows:
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    for cell in cells:
                        # Vérifie si le texte du numéro de facture est contenu dans la cellule, avec strip pour éviter les espaces en trop
                        if numero_facture.strip() in cell.text.strip():
                            self.logger.info(f"Ligne trouvée avec le numéro facture {cell.text}")
                            
                            # S'assurer que l'élément est visible et défiler jusqu'à l'élément si nécessaire
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
                            driver.execute_script("arguments[0].click();", row)
                            self.logger.info(f"Ligne avec le numéro facture {numero_facture} sélectionnée et cliquée avec succès.")
                            return True  # Contrat trouvé et traité avec succès

                except Exception as e:
                    self.logger.error(f"Erreur lors de la vérification des cellules dans la ligne : {e}")
                    continue

            # Aucun contrat trouvé, log et retourne False
            self.logger.error(f"La ligne avec le numéro facture {numero_facture} n'a pas été trouvée ou n'est pas cliquable.")
            return False

        except TimeoutException:
            self.logger.error(f"Le tableau des documents ou la ligne avec le numéro facture {numero_facture} n'a pas été trouvée.")
            return False
        except Exception as e:
            self.logger.error(f"Erreur lors de la sélection de la ligne avec le numéro facture {numero_facture} : {e}")
            return False

    def wait_for_modal(self, driver):
        """
        Attendre l'apparition de la modal pour poursuivre les opérations sur le contrat.
        """
        try:
            self.logger.info("Attente de l'apparition de la modal...")
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "body > div.bootbox.modal.fade.in > div > div > div"))
            )
            self.logger.info("Modal apparue et prête à être utilisée.")
        except TimeoutException:
            self.logger.error("La modal n'a pas été trouvée.")
        except Exception as e:
            self.logger.error(f"Erreur lors de l'attente de la modal : {e}")

    def verifier_SIRET(self, driver, SIRET):
        """
        Vérifie si le SIRET destinataire correspond au SIRET attendu.
        """
        try:
            siret_input = driver.find_element(By.ID, "m_client_siret")
            siret_valeur = siret_input.get_attribute("value")
            return siret_valeur == SIRET
        except Exception as e:
            self.logger.error(f"Erreur lors de la vérification du SIRET : {e}")
            return False

    def remplacer_siret(self, driver, siret_destinataire):
        """
        Remplace le SIRET par celui du payeur.
        """
        try:
            siret_input = driver.find_element(By.ID, "m_client_siret")
            siret_input.clear()
            siret_input.send_keys(siret_destinataire)
            time.sleep(2)
            
            self.logger.info(f"SIRET remplacé par {siret_destinataire}.")
        except Exception as e:
            self.logger.error(f"Erreur lors du remplacement du SIRET : {e}")

    def process_contract(self, driver, numero_facture, siret_destinataire, identifiant, mot_de_passe):
        """
        Traite un contrat spécifique avec Selenium.
        """
        wait = WebDriverWait(driver, 20)
        start_time = time.time()

        # Initialisation des variables critiques
        success = False
        message = "Erreur inattendue"
        duration = 0

        try:
            self.logger.info(f"Début du traitement du contrat {numero_facture}.")

            # Étape 1 : Clic sur "Rejets AIFE"
            self.click_rejets_aife(driver)

            # Étape 2 : Saisie du numéro de facture
            self.enter_num_facture(driver, numero_facture)

            # Étape 3 : Sélection de la ligne de la facture
            if not self.select_row_by_facture(driver, numero_facture):
                message = "Contrat non trouvé"
                self.logger.warning(f"Contrat {numero_facture} non trouvé. Passage au suivant.")
                return numero_facture, success, message, duration
            time.sleep(2)
            # Étape 4 : Attente de la modal
            self.wait_for_modal(driver)
            time.sleep(2)
            # Étape 5 : Remplacement du SIRET destinataire
            self.remplacer_siret(driver, siret_destinataire)
            time.sleep(2)
            # Étape 7 : Écriture d'un commentaire
            self.write_comment(driver, "Siret destinataire corrigé selon Prmedi")
            time.sleep(2)
            # Étape 6 : Clic sur le bouton "Sauvegarder"
            self.click_button_by_text(driver, "Sauvegarder")
            time.sleep(2)
            # Étape 8 : Gestion du bouton "Valider" et clic sur la modale
            self.click_and_validate_modal(driver, "Valider", "Valider")

            # Étape 10 : Vérification de la page d'erreur
            if self.is_error_page(driver):
                raise Exception("Page d'erreur détectée.")
            # Succès
            success = True
            message = "Succès"
        except Exception as e:
            message = str(e)
            self.log_error(numero_facture, message)
        finally:
            # Calcul de la durée et mise à jour des métriques
            duration = int(time.time() - start_time)
            self.update_metrics(success, duration, numero_facture, message)
            return numero_facture, success, message, duration



    def dictionnaire_siret(self, excel_path):
        """
        Lit un fichier Excel et construit un dictionnaire avec les numéros de contrat,
        SIRET destinataire, et SIRET payeur.
        """
        df = pd.read_excel(excel_path, dtype={'SIRET DESTINATAIRE': str})
        dictionnaire_siret = {}
        for _, row in df.iterrows():
            numero_facture = str(row['Contrat Nb'])
            siret = str(row['SIRET'])
            siret_destinataire = row['SIRET DESTINATAIRE']
            dictionnaire_siret[numero_facture] = {
                "SIRET": siret,
                "SIRET DESTINATAIRE": siret_destinataire
            }
        return dictionnaire_siret
    

    def process_contract_task(self, numero_facture, siret_destinataire, identifiant, mot_de_passe):
        """
        Traite un contrat spécifique dans un thread et met à jour les métriques.
        """
        driver = self.pool.get_driver()

        # Initialisation des variables critiques
        success = False
        message = "Erreur inattendue"
        duration = 0

        try:
            # Appel de la méthode principale de traitement
            numero_facture, success, message, duration = self.process_contract(
                driver, numero_facture, siret_destinataire, identifiant, mot_de_passe
            )
        except Exception as e:
            # Gestion des erreurs inattendues
            message = f"Erreur inattendue : {e}"
            self.log_error(numero_facture, message)
        finally:
            # Toujours retourner le driver au pool
            self.pool.return_driver(driver)

        return numero_facture, success, message, duration


    def main(self, excel_path):
        """
        Fonction principale pour gérer le traitement des contrats avec multi-threading.
        """
        self.logger.info("Démarrage du RPA Seres avec multithreading...")

        # Charger les identifiants depuis les variables d'environnement
        identifiant = os.getenv("IDENTIFIANT_SERES")
        mot_de_passe = os.getenv("MOT_DE_PASSE_SERES")
        if not identifiant or not mot_de_passe:
            self.logger.error("Identifiant ou mot de passe manquant. Arrêt du processus.")
            return

        # Préparer les données depuis le fichier Excel et JSON
        json_path = 'data/numeros_contrat_seres.json'
        extract_contrat_numbers_to_json(excel_path, json_path)
        dictionnaire_siret = self.dictionnaire_siret(excel_path)
        facture_numbers = self.process_json_files(json_path)

        # Utilisation de ThreadPoolExecutor pour le traitement multi-threading
        with ThreadPoolExecutor(max_workers=1) as executor:  # Ajustez max_workers selon les besoins
            futures = []
            for numero_facture in facture_numbers[:10]:
                # Récupération des informations SIRET pour chaque contrat
                siret_info = dictionnaire_siret.get(numero_facture)
                siret_destinataire = siret_info["SIRET DESTINATAIRE"] if siret_info else None

                if siret_destinataire:
                    # Création de threads pour traiter chaque contrat
                    future = executor.submit(
                        self.process_contract_task,
                        numero_facture,
                        siret_destinataire,
                        identifiant,
                        mot_de_passe
                    )
                    futures.append(future)

            # Récupération des résultats au fur et à mesure
            for future in as_completed(futures):
                try:
                    numero_facture, success, message, duration = future.result()
                    if success:
                        self.logger.info(f"Contrat {numero_facture} traité avec succès en {duration:.2f} secondes.")
                    else:
                        self.logger.warning(f"Échec du traitement du contrat {numero_facture} : {message}")
                except Exception as e:
                    self.logger.error(f"Erreur dans un thread de traitement : {e}")

        self.logger.info("Traitement de tous les contrats terminé.")

    def start(self, excel_path="data/data_traitement/Rejet SERES - Correction SIRET destinataire - 20241118.xlsx"):
        """
        Démarre le traitement du RPA Seres avec tableau de bord et logs.
        """
        self.logger.info(f"Démarrage du RPA Seres avec le fichier : {excel_path}")
        threading.Thread(target=self.start_dashboard).start()
        self.logger.info(f"Démarrage du tableau de board  dans http://127.0.0.1:8050")
          # Lancement du tableau de bord
        self.main(excel_path)  # Démarrage du traitement principal
