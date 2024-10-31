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
from concurrent.futures import ProcessPoolExecutor, as_completed

load_dotenv()
# Siret destinataire corrigé selon Prmedi
def process_contract_task(numero_facture, siret_destinataire, identifiant, mot_de_passe, url, pool):
            """
            Fonction pour traiter un contrat dans un processus séparé.
            """
            driver = None
            try:
                driver = pool.get_driver()  # Récupère un WebDriver pour ce contrat
                result = pool.process_contract(driver, numero_facture, siret_destinataire, identifiant, mot_de_passe)
                return result
            except Exception as e:
                print(f"Erreur lors du traitement du contrat {numero_facture} : {e}")
                return numero_facture, False, "Erreur", 0
            finally:
                if driver:
                    pool.return_driver(driver)
class SeresRPA:

    def __init__(self, pool, logger=None):
        """
        Initialise la classe avec un WebDriverPool, un logger et les identifiants.
        """
        self.pool = pool
        self.logger = logger or setup_logger("seres_case.log")
        self.url = "https://portail.e-facture.net/saml/saml-login.php?nomSP=ARTIMON_PROD"
        self.file_lock = threading.Lock()
    
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
    
    def click_sauvegarde_button(self, driver, numero_facture):
        """
        Essaie de cliquer sur le bouton de sauvegarde avec deux sélecteurs différents sans générer d'erreur si un sélecteur est introuvable.
        """
        selectors = [
            "#indexation-inner > div:nth-child(4) > button.btn.btn-primary",
            "#indexation-inner > div:nth-child(5) > button.btn.btn-primary"
        ]

        for selector in selectors:
            self.logger.debug(f"Essai du sélecteur : {selector}")
            
            # Utiliser find_elements pour éviter une erreur si le sélecteur ne correspond à aucun élément
            buttons = driver.find_elements(By.CSS_SELECTOR, selector)
            
            # Vérifier si le bouton est trouvé avant d'essayer de cliquer
            if buttons:
                validate_button = buttons[0]  # Prendre le premier bouton trouvé
                ActionChains(driver).move_to_element(validate_button).click().perform()
                self.logger.info("Bouton de sauvegarde cliqué avec succès.")
                return  # Sortie de la fonction si le clic est réussi

            else:
                self.logger.debug(f"Aucun bouton trouvé avec le sélecteur {selector}")

        # Si aucun des sélecteurs n'a fonctionné, log et enregistrer l'erreur
        self.logger.error(f"Échec du clic sur le bouton de sauvegarde pour le numéro facture {numero_facture} avec les sélecteurs fournis.")
        self.save_non_modifiable(numero_facture, "sauvegarde_button_erreur.json")


    def click_validate_button_modale(self, driver, numero_facture):
        try:
            time.sleep(5)
            # Localiser le bouton à l'aide du sélecteur CSS
            validate_button = driver.find_element(By.CSS_SELECTOR, "body > div.bootbox.modal.fade.bootbox-confirm.in > div > div > div.modal-footer > button.btn.btn-primary")
            # Scroller jusqu'au bouton si nécessaire et cliquer
            ActionChains(driver).move_to_element(validate_button).click().perform()
            self.logger.info("Bouton de validation cliqué avec succès.")
             # Attendre la potentielle apparition de l'alerte (timeout après 5 secondes)
            try:
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.alert.alert-danger.alert-dismissable.message"))
                )
                self.logger.error("Une erreur est survenue : Message d'alerte trouvé.")

                # Appeler la fonction save_non_modifiable pour sauvegarder le screenshot et le numéro de facture
                self.save_non_modifiable(numero_facture, "numero_SE_manquant.json")
            
            except TimeoutException:
                # Si l'erreur n'apparaît pas dans les 5 secondes, on considère que tout est ok
                self.logger.info("Pas d'erreur détectée après le clic.")
        except Exception as e:
            self.logger.error(f"Erreur lors du clic sur le bouton de validation Modale : {e}")
            self.save_non_modifiable(numero_facture, "validation_modale_button_erreur.json")


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
            # Attendre que le champ de saisie soit visible
            self.logger.info(f"Saisie du numéro facture : {numero_facture}")
            input_numfacture = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "gs_NUMFACTURE"))
            )

            # Effacer le champ et entrer le numéro de facture
            input_numfacture.clear()
            time.sleep(1)  # Attendre un court moment pour éviter d'éventuels conflits
            input_numfacture.send_keys(numero_facture)
            time.sleep(1)  # Attendre un court moment pour éviter d'éventuels conflits

            # Essayer d'appuyer sur Entrée et valider si la recherche s'est bien lancée
            input_numfacture.send_keys(Keys.ENTER)
            self.logger.debug("Appui sur Entrée pour lancer la recherche.")
            
            # Alternative : si ENTER ne fonctionne pas, essayer avec un clic sur un bouton de recherche
            try:
                search_button = driver.find_element(By.XPATH, "//button[@type='submit']")
                search_button.click()
                self.logger.debug("Clic sur le bouton de recherche en complément de l'appui sur Entrée.")
            except Exception:
                self.logger.warning("Le bouton de recherche n'a pas été trouvé, uniquement Entrée utilisé.")
            
            # Vérification : attendre un changement dans la table de résultats
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//table[contains(@id, 'list-documents')]//tr"))
            )
            
            self.logger.info(f"Numéro facture {numero_facture} saisi et recherche lancée avec succès.")
        
        except TimeoutException:
            self.logger.error("Le champ de saisie pour 'NUMFACTURE' n'a pas été trouvé.")
        except Exception as e:
            self.logger.error(f"Erreur lors de la saisie du numéro facture : {e}")
        finally:
            # Nettoyage du champ de saisie après le traitement
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
        wait = WebDriverWait(driver, 20)
        start_time = time.time()

        try:
            if not driver:
                raise Exception("Driver non initialisé")
            self.logger.info(f"Début du traitement du contrat {numero_facture}...")

            # Clic sur la div "Rejets AIFE"
            self.click_rejets_aife(driver)

            # Saisie et recherche du numéro de facture
            self.enter_num_facture(driver, numero_facture)

            # Sélection de la ligne de la facture - si non trouvée, arrêter et passer au contrat suivant
            if not self.select_row_by_facture(driver, numero_facture):
                self.logger.warning(f"Contrat {numero_facture} non trouvé. Passage au contrat suivant.")
                return numero_facture, False, "Contrat non trouvé", 0

            # Attente de la modal
            self.wait_for_modal(driver)

            # Remplacement du SIRET destinataire
            self.remplacer_siret(driver, siret_destinataire)
            time.sleep(2)

            # Clic sur le bouton de sauvegarde
            self.click_sauvegarde_button(driver, numero_facture)

            # Écriture d'un commentaire
            self.write_comment(driver, "Siret destinataire corrigé selon Prmedi")

            # Clic sur le bouton de validation
            self.click_validate_button(driver, numero_facture)
            time.sleep(3)

            # Clic sur le bouton de validation de la modale
            self.click_validate_button_modale(driver, numero_facture)

            # Vérifier la présence de la page d'erreur
            if self.is_error_page(driver):
                self.logger.warning("Page d'erreur détectée. Relance du processus...")
                driver.refresh()
                self.save_non_modifiable(numero_facture, "facture_error_cause_page_erreur.json")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_facture}: {e}")
            self.save_non_modifiable(numero_facture, "numero_facture_erreur.json")

            # Réinitialisation du driver après l'erreur
            try:
                driver.get(self.url)
            except Exception as reset_error:
                self.logger.error(f"Erreur lors de la réinitialisation du WebDriver pour {numero_facture}: {reset_error}")
                driver.quit()
                driver = None

        end_time = time.time()
        duration = int(end_time - start_time)
        return numero_facture, True, "Succès", duration



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
    

    def main(self, excel_path, pool):
        """
        Fonction principale pour traiter tous les contrats en parallèle.
        """
        logger = setup_logger("SeresRPA")
        logger.debug("Démarrage du RPA Seres avec multiprocessing...")

        # Charger les variables d'environnement au début
        identifiant = os.getenv("IDENTIFIANT_SERES")
        mot_de_passe = os.getenv("MOT_DE_PASSE_SERES")
        if not identifiant or not mot_de_passe:
            logger.error("Identifiant ou mot de passe manquant.")
            return

        # Extraire les contrats depuis le fichier
        json_path = 'data/numeros_contrat_seres.json'
        extract_contrat_numbers_to_json(excel_path, json_path)
        dictionnaire_siret = self.dictionnaire_siret(excel_path)
        facture_numbers = self.process_json_files(json_path)

        # Utilisation de ProcessPoolExecutor pour le traitement multi-processus
        with ProcessPoolExecutor() as executor:
            futures = []
            for numero_facture in facture_numbers:
                siret_info = dictionnaire_siret.get(numero_facture)
                siret_destinataire = siret_info["SIRET DESTINATAIRE"] if siret_info else None

                if siret_destinataire:
                    # Lancer chaque contrat dans un processus séparé sans `self`
                    future = executor.submit(
                        process_contract_task,
                        numero_facture,
                        siret_destinataire,
                        identifiant,
                        mot_de_passe,
                        self.url,  # URL de démarrage
                        self.pool
                    )
                    futures.append(future)

            # Collecter les résultats au fur et à mesure de l'achèvement des contrats
            for future in as_completed(futures):
                try:
                    numero_facture, result, contrat_type, duration = future.result()
                    if result:
                        logger.info(f"Contrat {numero_facture} traité avec succès en {duration} secondes.")
                    else:
                        logger.warning(f"Échec du traitement du contrat {numero_facture}.")
                except Exception as e:
                    logger.error(f"Erreur dans le processus de traitement : {e}")

        logger.info("Traitement de tous les contrats terminé.")


    def start(self, excel_path="data/data_traitement/Feuille de traitement problème SIRET - Rejets SERES.xlsx"):
        """
        Démarre le traitement du RPA Seres.
        """
        self.logger.info(f"Démarrage du RPA Seres avec le fichier {excel_path}")
        self.main(excel_path, self.pool)  # Correction ici
