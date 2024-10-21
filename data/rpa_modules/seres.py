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

from rpa_modules.debug import setup_logger
from selenium.webdriver.common.action_chains import ActionChains
from rpa_modules.data_processing import extract_contrat_numbers_to_json

load_dotenv()
# Siret destinataire corrigé selon Prmedi
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

            input_identifiant = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#login")))
            input_identifiant.clear()
            input_identifiant.send_keys(identifiant)

            input_mot_de_passe = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#acct_pass")))
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

            # Prendre et sauvegarder le screenshot
            driver.save_screenshot(screenshot_filename)
            print(f"Screenshot sauvegardé sous {screenshot_filename}")

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
    
    def click_cookie_consent_button(self, driver):
        try:
            # Localiser le bouton à l'aide du sélecteur CSS
            consent_button = driver.find_element(By.CSS_SELECTOR, "#cookieConsentContainer > a")
            # Scroller jusqu'au bouton si nécessaire et cliquer
            ActionChains(driver).move_to_element(consent_button).click().perform()
            self.logger.info("Bouton de consentement cliqué avec succès.")
        except Exception as e:
            self.logger.error(f"Erreur lors du clic sur le bouton : {e}")
    
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
    
    def click_validate_button(self, driver):
        try:
            # Localiser le bouton à l'aide du sélecteur CSS
            validate_button = driver.find_element(By.CSS_SELECTOR, "#validate")
            # Scroller jusqu'au bouton si nécessaire et cliquer
            ActionChains(driver).move_to_element(validate_button).click().perform()
            self.logger.info("Bouton de validation cliqué avec succès.")
        except Exception as e:
            self.logger.error(f"Erreur lors du clic sur le bouton de validation Valider: {e}")
    
    def click_validate_button_modale(self, driver, numero_facture):
        try:
            # Localiser le bouton à l'aide du sélecteur CSS
            validate_button = driver.find_element(By.CSS_SELECTOR, "body > div.bootbox.modal.fade.bootbox-confirm.in > div > div > div.modal-footer > button.btn.btn-primary")
            # Scroller jusqu'au bouton si nécessaire et cliquer
            ActionChains(driver).move_to_element(validate_button).click().perform()
            self.logger.info("Bouton de validation cliqué avec succès.")
             # Attendre la potentielle apparition de l'alerte (timeout après 5 secondes)
            try:
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.alert.alert-danger.alert-dismissable.message"))
                )
                self.logger.error("Une erreur est survenue : Message d'alerte trouvé.")

                # Appeler la fonction save_non_modifiable pour sauvegarder le screenshot et le numéro de facture
                self.save_non_modifiable(driver, numero_facture, "numero_facture_erreur.json")
            
            except TimeoutException:
                # Si l'erreur n'apparaît pas dans les 5 secondes, on considère que tout est ok
                self.logger.info("Pas d'erreur détectée après le clic.")
        except Exception as e:
            self.logger.error(f"Erreur lors du clic sur le bouton de validation Modale : {e}")

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
        Saisit le numéro de facture dans le champ de recherche et lance la recherche.
        """
        try:
            self.logger.info(f"Saisie du numéro facture : {numero_facture}")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "gs_NUMFACTURE"))
            )
            input_numfacture = driver.find_element(By.ID, "gs_NUMFACTURE")
            input_numfacture.send_keys(numero_facture)
            time.sleep(2)
            input_numfacture.send_keys(Keys.ENTER)
            time.sleep(2)
            self.logger.info(f"Numéro facture {numero_facture} saisi et recherche lancée.")
        except TimeoutException:
            self.logger.error("Le champ de saisie pour 'NUMFACTURE' n'a pas été trouvé.")
        except Exception as e:
            self.logger.error(f"Erreur lors de la saisie du numéro facture : {e}")

    def select_row_by_facture(self, driver, numero_facture):
        """
        Sélectionne la ligne du tableau correspondant au numéro facture fourni et clique dessus.
        :param driver: L'instance WebDriver.
        :param numero_facture: Le numéro facture à rechercher dans le tableau.
        """
        try:
            self.logger.info(f"Sélection de la ligne avec le numéro facture : {numero_facture}...")

            # Attendre que le tableau soit visible
            wait = WebDriverWait(driver, 20)

            # Attendre un peu après la saisie du numéro facture pour que le tableau se mette à jour
            time.sleep(5)

            # Rechercher le tableau contenant les documents
            table = wait.until(EC.presence_of_element_located((By.XPATH, "//table[contains(@id, 'list-documents')]")))

            # Récupérer toutes les lignes du tableau (exclure la première ligne si elle est un header)
            rows = table.find_elements(By.XPATH, ".//tr[not(contains(@class, 'jqgfirstrow'))]")
            self.logger.debug(f"Nombre de lignes trouvées dans le tableau : {len(rows)}")

            # Comparaison pour vérifier que le numéro facture est contenu dans la cellule (en ignorant les zéros)
            for row in rows:
                try:
                    # On tente de récupérer la cellule qui contient le numéro facture
                    cells = row.find_elements(By.TAG_NAME, "td")

                    # Boucle sur chaque cellule pour voir si le numéro facture est présent dans l'une d'elles
                    for cell in cells:
                        if numero_facture in cell.text:
                            self.logger.info(f"Ligne trouvée avec le numéro facture {cell.text}")

                            # Faire défiler jusqu'à l'élément si nécessaire
                            driver.execute_script("arguments[0].scrollIntoView(true);", row)

                            # Utiliser un clic JavaScript pour contourner les éventuels problèmes d'interaction
                            driver.execute_script("arguments[0].click();", row)

                            self.logger.info(f"Ligne avec le numéro facture {numero_facture} sélectionnée et cliquée avec succès.")
                            return

                except Exception as e:
                    self.logger.error(f"Erreur lors de la vérification des cellules dans la ligne : {e}")
                    continue

            # Si aucune ligne correspondante n'est trouvée
            self.logger.error(f"La ligne avec le numéro facture {numero_facture} n'a pas été trouvée ou n'est pas cliquable.")
            driver.save_screenshot(f"screenshot_error_{numero_facture}.png")

        except TimeoutException:
            self.logger.error(f"La ligne avec le numéro facture {numero_facture} n'a pas été trouvée ou n'est pas cliquable.")
            driver.save_screenshot(f"screenshot_error_{numero_facture}.png")
        except Exception as e:
            self.logger.error(f"Erreur lors de la sélection de la ligne avec le numéro facture {numero_facture} : {e}")
            driver.save_screenshot(f"screenshot_error_general_{numero_facture}.png")


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

    def verifier_siret_destinataire(self, driver, siret_destinataire):
        """
        Vérifie si le SIRET destinataire correspond au SIRET attendu.
        """
        try:
            siret_input = driver.find_element(By.ID, "m_client_siret")
            siret_valeur = siret_input.get_attribute("value")
            return siret_valeur == siret_destinataire
        except Exception as e:
            self.logger.error(f"Erreur lors de la vérification du SIRET : {e}")
            return False

    def remplacer_siret(self, driver, siret_payeur):
        """
        Remplace le SIRET par celui du payeur.
        """
        try:
            siret_input = driver.find_element(By.ID, "m_client_siret")
            siret_input.clear()
            siret_input.send_keys(siret_payeur)
            time.sleep(2)
            
            self.logger.info(f"SIRET remplacé par {siret_payeur}.")
        except Exception as e:
            self.logger.error(f"Erreur lors du remplacement du SIRET : {e}")

    def process_contract(self, numero_facture, siret_destinataire, siret_payeur, driver, identifiant, mot_de_passe):
        wait = WebDriverWait(driver, 10)
        retry_count = 0
        max_retries = 3

        while retry_count < max_retries:
            try:
                self.logger.info(f"Début du traitement du contrat {numero_facture}...")

                # self.login(driver, wait, identifiant, mot_de_passe)
                self.click_rejets_aife(driver)
                self.enter_num_facture(driver, numero_facture)
                self.select_row_by_facture(driver, numero_facture)
                self.wait_for_modal(driver)


                self.click_cookie_consent_button(driver)
                time.sleep(3)

                # if self.verifier_siret_destinataire(driver, siret_destinataire):
                self.remplacer_siret(driver, siret_payeur)
                time.sleep(2)
                self.write_comment(driver, "Siret destinataire corrigé selon Prmedi")
                self.click_validate_button(driver)
                time.sleep(2)
                self.click_validate_button_modale(driver, numero_facture)
                # Vérifier si la page d'erreur est présente
                if self.is_error_page(driver):
                    self.logger.warning("Page d'erreur détectée. Relance du processus...")
                    retry_count += 1
                    driver.refresh()  # Recharger la page ou redémarrer le processus
                    continue  # Recommencer le processus

                self.logger.info(f"Contrat {numero_facture} traité avec succès.")
                break  # Sortir de la boucle si le traitement est réussi

            except Exception as e:
                self.logger.error(f"Erreur lors du traitement du contrat {numero_facture}: {e}")
                retry_count += 1
                driver.refresh()  # Recharger la page en cas d'erreur

        if retry_count == max_retries:
            self.logger.error(f"Le processus pour le contrat {numero_facture} a échoué après {max_retries} tentatives.")


    def dictionnaire_siret(self, excel_path):
        """
        Lit un fichier Excel et construit un dictionnaire avec les numéros de contrat,
        SIRET destinataire, et SIRET payeur.
        """
        df = pd.read_excel(excel_path)
        dictionnaire_siret = {}
        for _, row in df.iterrows():
            numero_facture = str(row['Numéro de facture'])
            siret_destinataire = str(row['SIRET emetteur'])
            siret_payeur = str(row['SIRET destinataire'])
            dictionnaire_siret[numero_facture] = {
                "siret_destinataire": siret_destinataire,
                "siret_payeur": siret_payeur
            }
        return dictionnaire_siret

    def main(self, excel_path):
        self.logger.debug("Démarrage du RPA Seres...")
        json_path = 'data/numeros_contrat_seres.json'
        extract_contrat_numbers_to_json(excel_path, json_path)

        dictionnaire_siret = self.dictionnaire_siret(excel_path)
        facture_numbers = self.process_json_files(json_path)

        identifiant = os.getenv("IDENTIFIANT_SERES")
        mot_de_passe = os.getenv("MOT_DE_PASSE_SERES")

        if not identifiant or not mot_de_passe:
            self.logger.error("Identifiant ou mot de passe manquant.")
            return

        try:
            for numero_facture in facture_numbers:
                numero_facture = str(numero_facture)
                siret_info = dictionnaire_siret.get(numero_facture)

                if siret_info:
                    siret_destinataire = siret_info["siret_destinataire"]
                    siret_payeur = siret_info["siret_payeur"]

                    driver = self.pool.get_driver(self.url)  # Obtenir un WebDriver pour traiter ce contrat
                    self.process_contract(numero_facture, siret_destinataire, siret_payeur, driver, identifiant, mot_de_passe)
                    self.pool.return_driver(driver)  # Retourner le WebDriver au pool après utilisation
                else:
                    self.logger.error(f"Numéro de contrat {numero_facture} introuvable.")
        finally:
            self.pool.close_all()

    def start(self, excel_path="data/data_traitement/Feuille de traitement problème SIRET - Rejets SERES.xlsx"):
        """
        Démarre le traitement du RPA Seres.
        """
        self.logger.info(f"Démarrage du RPA Seres avec le fichier {excel_path}")
        self.main(excel_path)
