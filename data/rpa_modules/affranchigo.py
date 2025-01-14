import os
import json
import copy
import traceback
import pandas as pd
import threading
import time
import psutil
import csv
from dotenv import load_dotenv
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor
from selenium.common.exceptions import TimeoutException, NoSuchElementException
# Importation du setup_logger uniquement
from concurrent.futures import ThreadPoolExecutor, as_completed
from rpa_modules.debug import setup_logger
from rpa_modules.data_processing import extract_contrat_numbers_to_json

load_dotenv()

STOP_FLAG = False

class AffranchigoRPA:
    def __init__(self, pool, logger=None):
        """
        Classe RPA pour Affranchigo.
        :param pool: Pool de WebDrivers partagé.
        :param logger: Logger pour suivre les logs.
        """
        self.pool = pool
        self.logger = logger or setup_logger('affranchigo.log')
        self.file_lock = threading.Lock()
        self.submit_lock = threading.Lock()
        self.STOP_FLAG = False

        # URL spécifique à Affranchigo
        self.url = "https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/index/acces-dc?profil=ADV"

    def terminate_high_cpu_processes(self, process_name="msedge.exe", cpu_threshold=15, max_runtime=300):
        for proc in psutil.process_iter(['pid', 'name', 'cpu_percent', 'create_time']):
            try:
                if proc.info['name'] == process_name:
                    cpu_usage = proc.cpu_percent(interval=1)
                    runtime = time.time() - proc.create_time()

                    if cpu_usage > cpu_threshold and runtime > max_runtime:
                        ppid = proc.ppid()
                        if not psutil.pid_exists(ppid):
                            self.logger.warning(
                                f"Terminating {process_name} with PID {proc.info['pid']} due to high CPU usage: {cpu_usage}% and runtime: {runtime} seconds")
                            proc.terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
    
    def wait_for_page_load(self, driver, timeout=20):
        try:
            WebDriverWait(driver, timeout).until(
                lambda d: d.execute_script('return document.readyState') == 'complete'
            )
            self.logger.debug("La page est complètement chargée.")
        except TimeoutException as e:
            self.logger.error(f"Erreur Timeout lors du chargement de la page : {e}")
            driver.save_screenshot("page_load_timeout.png")



    def save_results_to_csv(self, results, contrat_types_count, multisites_count, non_modifiables_count, csv_file_path):
        if not results:
            self.logger.error("Aucun résultat à enregistrer.")
            return

        headers = ['Numéro de contrat', 'Résultat', 'Durée (secondes)', 'Type de contrat']
        df = pd.DataFrame(results, columns=headers)
        df['Durée (secondes)'] = df['Durée (secondes)'].astype(int)

        df.to_csv(csv_file_path, index=False)

        types_count_df = pd.DataFrame(list(contrat_types_count.items()), columns=['Type de contrat', 'Nombre'])
        types_count_df['Nombre'] = types_count_df['Nombre'].astype(int)
        types_count_df.to_csv('contract_types_count.csv', index=False)

        with open(csv_file_path, 'a') as f:
            f.write(f"\nNombre de contrats multisites: {multisites_count}")
            f.write(f"\nNombre de contrats non modifiables: {non_modifiables_count}")

        self.logger.info(f"Résultats enregistrés dans {csv_file_path} et décompte des types de contrats dans contract_types_count.csv")

    def login(self, driver, wait, identifiant, mot_de_passe):
        self.logger.debug("Tentative de connexion...")
        try:
            input_identifiant = wait.until(EC.presence_of_element_located((By.ID, "AUTHENTICATION.LOGIN")))
            input_identifiant.clear()
            input_identifiant.send_keys(identifiant)
            input_identifiant.send_keys(Keys.RETURN)

            input_mot_de_passe = wait.until(EC.presence_of_element_located((By.ID, "AUTHENTICATION.PASSWORD")))
            input_mot_de_passe.clear()
            input_mot_de_passe.send_keys(mot_de_passe)
            input_mot_de_passe.send_keys(Keys.RETURN)
        except TimeoutException:
            self.logger.info("Dejà connecté ou le champ d'identifiant n'est pas présent.")
        except Exception as e:
            self.logger.error( "Problème Login")

    def process_json_files(self, file_path):
        self.logger.debug(f"Traitement du fichier JSON pour les contrats à partir du fichier : {file_path}")
        numeros_contrat = []

        if os.path.exists(file_path):
            with open(file_path, "r") as file:
                try:
                    data = json.load(file)
                    numeros_contrat = list(data.values())
                    self.logger.debug(f"Numéros de contrat récupérés : {numeros_contrat}")
                except json.JSONDecodeError as e:
                    self.logger.error(f"Erreur lors du chargement du fichier JSON : {e}")
        else:
            self.logger.error(f"Le fichier JSON {file_path} est introuvable.")

        return numeros_contrat

    def submit_contract_number(self, driver, wait, numero):
        with self.submit_lock:
            try:
                self.logger.info(f"Soumission du numéro de contrat {numero}...")

                input_element = wait.until(EC.element_to_be_clickable((By.ID, "idContrat")))
                self.logger.debug("Champ de texte trouvé et cliquable.")
                input_element.clear()
                input_element.send_keys(numero)

                submit_button = wait.until(EC.element_to_be_clickable((By.ID, "btnSubmitContrat_accesRDC")))
                self.logger.debug("Bouton de soumission trouvé et cliquable.")
                submit_button.click()

                self.wait_for_page_load(driver)

                self.logger.info(f"Numéro de contrat {numero} soumis avec succès.")
            except TimeoutException as e:
                self.logger.warning(f"Timeout lors de la soumission du contrat {numero}, la page pourrait ne pas avoir réagi à temps.")
                driver.save_screenshot(f"timeout_error_{numero}.png")
            except NoSuchElementException as e:
                self.logger.error(f"Élément introuvable pour la soumission du contrat {numero}: {e}")
                driver.save_screenshot(f"no_element_error_{numero}.png")
            except Exception as e:
                self.logger.error(f"Erreur inattendue lors de la soumission du contrat {numero}: {e}")
                driver.save_screenshot(f"general_error_{numero}.png")

    def save_processed_contracts(self, contrats, file_path="numeros_contrat_traites.json"):
        try:
            with self.file_lock, open(file_path, "r+") as file:
                try:
                    existing_data = json.load(file)
                except json.JSONDecodeError:
                    existing_data = []
                updated_data = existing_data + [c for c in contrats if c not in existing_data]
                file.seek(0)
                file.truncate()
                json.dump(updated_data, file)
        except FileNotFoundError:
            with self.file_lock, open(file_path, "w") as file:
                json.dump(contrats, file)

    def handle_driver_cleanup(self, driver, numero_contrat):
        """
        Gère le nettoyage, la réinitialisation et le retour du WebDriver au pool.
        """
        if not driver:
            return

        try:
            # Réinitialiser le WebDriver à l'URL de départ
            driver.get(self.url)
            self.logger.debug(f"WebDriver réinitialisé à l'URL de départ pour le contrat {numero_contrat}.")
            self.pool.return_driver(driver)
            self.logger.debug(f"WebDriver retourné au pool pour le contrat {numero_contrat}.")
        except Exception as reset_error:
            self.logger.error(f"Erreur lors de la réinitialisation ou du retour du WebDriver au pool pour {numero_contrat}: {reset_error}")
            try:
                driver.quit()
                self.logger.debug("WebDriver défectueux fermé avec succès.")
            except Exception as quit_error:
                self.logger.error(f"Erreur lors de la fermeture du WebDriver: {quit_error}")
            finally:
                try:
                    # Créer un nouveau WebDriver si le précédent a échoué
                    new_driver = self.pool.create_driver()
                    self.pool.return_driver(new_driver)
                    self.logger.debug("Nouveau WebDriver créé et ajouté au pool après défaillance.")
                except Exception as creation_error:
                    self.logger.critical(f"Erreur critique lors de la création d'un nouveau WebDriver: {creation_error}")
                    
    def create_dictionnaire(self, excel_path):
        df = pd.read_excel(excel_path)
        dictionnaire = {}
        for index, row in df.iterrows():
            numero_contrat = row['Contrat Nb']
            regate_depot_ancien = str(int(row['Code REGATE Dépôt actuel'])) if not pd.isna(row['Code REGATE Dépôt actuel']) else None
            regate_depot_nouveau = str(int(row['Nouveau code REGATE Dépôt'])) if not pd.isna(row['Nouveau code REGATE Dépôt']) else None
            regate_traitement_ancien = str(int(row['Code REGATE Traitement actuel'])) if not pd.isna(row['Code REGATE Traitement actuel']) else None
            regate_traitement_nouveau = str(int(row['Nouveau code REGATE Traitement'])) if not pd.isna(row['Nouveau code REGATE Traitement']) else None

            dictionnaire[numero_contrat] = {
                'Code REGATE Dépôt actuel': regate_depot_ancien,
                'Nouveau code REGATE Dépôt': regate_depot_nouveau,
                'Code REGATE Traitement actuel': regate_traitement_ancien,
                'Nouveau code REGATE Traitement': regate_traitement_nouveau
            }
        return dictionnaire


    def save_non_modifiable_contract_mutli_sites(self, contrat_number, file_path="annexe_multisites.json"):
        data = set()
        try:
            with self.file_lock, open(file_path, "r") as file:
                existing_data = json.load(file)
                data.update(existing_data)
        except FileNotFoundError:
            self.logger.debug("Le fichier n'existe pas, il sera créé.")
        except json.JSONDecodeError:
            self.logger.error("Erreur de décodage JSON, le fichier est peut-être corrompu.")

        if contrat_number not in data:
            data.add(contrat_number)
            with self.file_lock, open(file_path, "w") as file:
                json.dump(list(data), file)
            self.logger.debug(f"Contrat numéro {contrat_number} ajouté.")
            return True
        else:
            self.logger.debug(f"Contrat numéro {contrat_number} déjà présent, non ajouté.")
            return False

    def save_non_modifiable(self, contrat_number, file_path="problèmes_contrats_2.json"):
        data = set()
        try:
            with self.file_lock, open(file_path, "r") as file:
                existing_data = json.load(file)
                data.update(existing_data)
        except FileNotFoundError:
            self.logger.debug("Le fichier n'existe pas, il sera créé.")
        except json.JSONDecodeError:
            self.logger.error("Erreur de décodage JSON, le fichier est peut-être corrompu.")

        if contrat_number not in data:
            data.add(contrat_number)
            with self.file_lock, open(file_path, "w") as file:
                json.dump(list(data), file)
            self.logger.info(f"Contrat numéro {contrat_number} ajouté aux non modifiables.")
            return True
        else:
            self.logger.debug(f"Contrat numéro {contrat_number} déjà présent, non ajouté.")
            return False

    def handle_non_clickable_element(self, driver, numero_contrat):
        """Gestion des contrats multi-sites ou cas spécifiques."""
        try:
            span_element = driver.find_element(By.ID, "detailsCategorieV")
            if span_element.text == "Annexe Multisites":
                self.save_non_modifiable_contract_mutli_sites(numero_contrat)
                self.logger.info(f"Contrat {numero_contrat} enregistré comme non modifiable (Annexe Multisites).")
            else:
                self.logger.debug(f"Contrat {numero_contrat} non marqué comme 'Annexe Multisites'. Texte détecté : {span_element.text}")
        
        except NoSuchElementException:
            self.logger.debug(f"L'élément 'detailsCategorieV' introuvable pour le contrat {numero_contrat}.")
            self.handle_driver_cleanup(driver, numero_contrat)
        
        except Exception as e:
            self.logger.error(f"Erreur lors de la gestion des contrats multi-sites pour {numero_contrat} : {e}")
            self.handle_driver_cleanup(driver, numero_contrat)
        


    def switch_to_iframe_and_click_modification(self, driver, wait, contrat_number):
        """
        Change vers un iframe et clique sur un bouton de modification.

        :param driver: Instance WebDriver.
        :param wait: Instance WebDriverWait.
        :param contrat_number: Numéro du contrat en cours.
        """
        self.logger.info(f"{contrat_number} * Changement vers iframe et tentative de clic sur 'Modification'...")
        iframe_selector = "#modalRefContrat > div > div > div.modal-body > iframe"
        modification_button_selector = "//permission/a[@href='#amendment']//div[@id='detailsModificationButton']"

        try:
            # Passer à l'iframe
            iframe = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, iframe_selector)))
            driver.switch_to.frame(iframe)
            self.logger.debug(f"{contrat_number} * Passé à l'iframe avec succès.")
            time.sleep(2)
            # Cliquer sur le bouton de modification
            bouton_modification = wait.until(EC.element_to_be_clickable((By.XPATH, modification_button_selector)))
            driver.execute_script("arguments[0].scrollIntoView(true);", bouton_modification)
            driver.execute_script("arguments[0].click();", bouton_modification)
            self.logger.info(f"{contrat_number} * Clic sur le bouton de modification effectué avec succès.")
        except Exception as e:
            self.logger.error(f"{contrat_number} * Erreur lors du traitement de l'iframe : {e}")
            self.handle_non_clickable_element(driver, contrat_number)
        finally:
            # Toujours revenir au contenu principal
            driver.switch_to.default_content()
            self.logger.debug(f"{contrat_number} * Retour au contenu principal effectué.")

    def wait_for_complete_redirection(self, driver, wait, numero_contrat, timeout=20):
        """
        Attend la redirection complète et clique sur un élément cible.

        :param driver: Instance WebDriver.
        :param wait: Instance WebDriverWait.
        :param numero_contrat: Numéro du contrat en cours.
        :param timeout: Temps maximal d'attente pour le chargement de l'élément cible.
        """
        self.logger.debug(f"{numero_contrat} * Attente de la redirection...")
        target_selector = "#content_offre > ul > li:nth-child(2) > a"

        try:
            # Cliquer sur l'élément cible
            element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, target_selector)))
            element.click()
            self.logger.info(f"{numero_contrat} * L'élément cible a été cliqué avec succès.")
        except Exception as e:
            self.logger.error(f"{numero_contrat} * Erreur lors de la redirection ou du clic sur l'élément cible : {e}")
            self.save_non_modifiable(numero_contrat)
        finally:
            self.logger.debug(f"{numero_contrat} * Fin de la tentative de redirection.")

    def modifications_conditions_ventes(self, driver, wait, numero_contrat, dictionnaire, dictionnaire_original):
        # Importer les modules localement pour éviter les boucles d'importation circulaire
        from rpa_modules.affranchigo_forfait_case import AffranchigoForfaitCase
        from rpa_modules.affranchigo_lib_case import AffranchigoLibCase
        from rpa_modules.destineo_case import DestineoCase
        from rpa_modules.frequenceo_case import FrequenceoCase
        from rpa_modules.proxicompte_case import ProxicompteCase
        from rpa_modules.collecte_remise_case import CollecteRemise
        from rpa_modules.affranchigo_premium import AffranchigoPremiumCase

        affranchigo_premium_case = AffranchigoPremiumCase(driver, self.logger)
        affranchigo_forfait_case = AffranchigoForfaitCase(driver, self.logger)
        affranchigo_liberte_case = AffranchigoLibCase(driver, self.logger)
        destineo_case = DestineoCase(driver, self.logger)
        frequenceo_case = FrequenceoCase(driver, self.logger)
        proxicompte_case = ProxicompteCase(driver, self.logger)
        collecte_remise_case = CollecteRemise(driver, self.logger)

        try:
            # Attendre que la page soit complètement chargée
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            self.logger.debug("La page est complètement chargée")

            # Attendre que la page soit complètement chargée avant de chercher l'élément
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
            time.sleep(3)
            h1_element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#header_offre_descr > h1")))

            h1_text = h1_element.text
            self.logger.info(f"Texte de l'en-tête H1: {h1_text}")
            if "Affranchigo Premium" in h1_text:
                self.logger.info("Contrat Affranchigo Premium")
                affranchigo_premium_case.handle_case_premium(numero_contrat, dictionnaire)
                return "Affranchigo liberté"
            elif "Affranchigo forfait" in h1_text:
                self.logger.info("Contrat Affranchigo forfait")
                if not isinstance(dictionnaire, dict):
                    self.logger.error(f"Type incorrect de 'dictionnaire' détecté: {type(dictionnaire)}; restauration en cours.")
                    dictionnaire = copy.deepcopy(dictionnaire_original)
                    self.logger.debug(f"Dictionnaire type: {type(dictionnaire)}")
                else:
                    affranchigo_forfait_case.handle_case_forfait(driver, numero_contrat, dictionnaire)
                return "Affranchigo forfait"
            elif "Affranchigo liberté" in h1_text:
                self.logger.info("Contrat Affranchigo liberté")
                affranchigo_liberte_case.handle_case_lib(numero_contrat, dictionnaire)
                return "Affranchigo liberté"
            elif "Destineo esprit libre" in h1_text:
                self.logger.info("Contrat Destineo")
                destineo_case.handle_case_destineo(driver, wait, numero_contrat, dictionnaire)
                return "Destineo esprit libre"
            elif "Frequenceo" in h1_text:
                self.logger.info("Contrat Frequenceo")
                frequenceo_case.handle_case_frequenceo(driver, wait, numero_contrat, dictionnaire)
                return "Frequenceo"
            elif "Proxicompte" in h1_text:
                proxicompte_case.handle_case_proxicompte(driver, numero_contrat, dictionnaire)
                self.logger.info("Contrat Proxicompte")
                return "Proxicompte"
            elif "Collecte Remise Plus" in h1_text:
                collecte_remise_case.handle_case_collecte_remise(driver, wait, numero_contrat, dictionnaire)
                self.logger.info("Contrat Collecte Remise Plus")
                return "Collecte Remise Plus"
            elif "Collecte et remise" in h1_text:
                collecte_remise_case.handle_case_collecte_remise(driver, wait, numero_contrat, dictionnaire)
                self.logger.info("Contrat Collecte et Remise")
                return "Collecte et remise"
            else:
                return "Inconnu"
        except Exception as e:
            self.logger.exception(f"Service non reconnu : {e}")
            self.handle_driver_cleanup(driver, numero_contrat)
            return "Erreur"

    def process_contract(self, driver, numero_contrat, dictionnaire, dictionnaire_original, identifiant, mot_de_passe):
        wait = WebDriverWait(driver, 20)  # Augmentation du délai pour le WebDriver
        start_time = time.time()

        try:
            if not driver:
                raise Exception("Driver non initialisé")

            # Traitement du contrat (soumission, modifications, etc.)
            self.submit_contract_number(driver, wait, numero_contrat)
            self.switch_to_iframe_and_click_modification(driver, wait, numero_contrat)
            time.sleep(3)
            self.wait_for_complete_redirection(driver, wait, numero_contrat)
            contrat_type = self.modifications_conditions_ventes(driver, wait, numero_contrat, dictionnaire, dictionnaire_original)

            self.logger.info(f"{numero_contrat} * Traitement terminé.")
            self.save_processed_contracts([numero_contrat])

            end_time = time.time()
            duration = int(end_time - start_time)

            return (numero_contrat, True, contrat_type, duration)

        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat}: {e}")
            self.save_non_modifiable(numero_contrat)

            # Réinitialisation du driver après l'erreur
            try:
                driver.get(self.url)
            except Exception as reset_error:
                self.logger.error(f"Erreur lors de la réinitialisation du WebDriver pour {numero_contrat}: {reset_error}")
                driver.quit()  # Fermer le driver si la réinitialisation échoue
                driver = None  # Forcer la recréation d'un nouveau driver pour le prochain contrat

            end_time = time.time()
            duration = int(end_time - start_time)
            return (numero_contrat, False, "Erreur", duration)


    def process_single_contract(self, numero_contrat, dictionnaire, dictionnaire_original, identifiant, mot_de_passe):
        """
        Fonction qui traite un contrat individuel dans un thread séparé.
        """
        driver = None
        try:
            driver = self.pool.get_driver()  # Récupère un WebDriver pour le contrat
            self.logger.debug(f"WebDriver récupéré avec succès pour le contrat {numero_contrat}.")

            # Appel à la fonction process_contract pour traiter le contrat
            return self.process_contract(driver, numero_contrat, dictionnaire, dictionnaire_original, identifiant, mot_de_passe)

        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat}: {e}", exc_info=True)
            return (numero_contrat, False, "Erreur", 0)

        finally:
            if driver:
                try:
                    # Réinitialiser le WebDriver à l'URL de départ
                    driver.get(self.url)
                    self.logger.debug(f"WebDriver réinitialisé à l'URL de départ pour le contrat {numero_contrat}.")
                except Exception as reset_error:
                    # Si la réinitialisation échoue, détruire le driver
                    self.logger.error(f"Erreur lors de la réinitialisation du WebDriver pour {numero_contrat}: {reset_error}")
                    try:
                        driver.quit()
                        self.logger.debug("WebDriver défectueux fermé avec succès.")
                    except Exception as quit_error:
                        self.logger.error(f"Erreur lors de la fermeture du WebDriver: {quit_error}")
                    finally:
                        driver = None  # Forcer la création d'un nouveau WebDriver

                # Retourner le WebDriver au pool ou en créer un nouveau
                if driver:
                    try:
                        self.pool.return_driver(driver)
                        self.logger.debug(f"WebDriver retourné au pool pour le contrat {numero_contrat}.")
                    except Exception as pool_error:
                        self.logger.error(f"Erreur lors du retour du WebDriver au pool pour {numero_contrat}: {pool_error}")
                        try:
                            driver.quit()
                            self.logger.debug("WebDriver fermé après échec de retour au pool.")
                        except Exception as quit_error:
                            self.logger.error(f"Erreur lors de la fermeture du WebDriver: {quit_error}")
                        finally:
                            driver = None
                else:
                    # Créer un nouveau WebDriver si le précédent a échoué
                    try:
                        new_driver = self.pool.create_driver()
                        self.pool.return_driver(new_driver)
                        self.logger.debug("Nouveau WebDriver créé et ajouté au pool après défaillance.")
                    except Exception as creation_error:
                        self.logger.critical(f"Erreur critique lors de la création d'un nouveau WebDriver: {creation_error}")


    def main(self, progress_callback=None, max_workers=5):
        """
        Méthode principale pour le traitement du RPA avec multi-threading.
        """
        self.logger.debug("Démarrage du RPA Affranchigo en multi-threading...")
        excel_path = "data/data_traitement/DIGNE PPDC - Transfert des contrats Affranchigo 130125.xlsx"
        json_path = 'data/numeros_contrat_robot.json'

        # Extraction des numéros de contrat
        extract_contrat_numbers_to_json(excel_path, json_path)

        dictionnaire_original = self.create_dictionnaire(excel_path)
        dictionnaire = copy.deepcopy(dictionnaire_original)

        contract_numbers = self.process_json_files(json_path)

        identifiant = os.getenv("IDENTIFIANT")
        mot_de_passe = os.getenv("MOT_DE_PASSE")

        # Variables pour collecter les statistiques
        results = []
        contrat_types_count = {}
        multisites_count = 0
        non_modifiables_count = 0

        # Utilisation d'un ThreadPoolExecutor pour le traitement multi-threading
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = []
            for numero_contrat in contract_numbers:
                self.logger.debug(f"Préparation pour traiter le contrat suivant: {numero_contrat}")

                # Planification du traitement de chaque contrat dans un thread séparé
                future = executor.submit(self.process_single_contract, numero_contrat, dictionnaire, dictionnaire_original, identifiant, mot_de_passe)
                futures.append(future)

            # Collecter les résultats des threads au fur et à mesure
            for future in as_completed(futures):
                try:
                    numero_contrat, result, contrat_type, duration = future.result()
                    results.append((numero_contrat, result, duration, contrat_type))

                    # Mettre à jour les statistiques
                    if contrat_type not in contrat_types_count:
                        contrat_types_count[contrat_type] = 0
                    contrat_types_count[contrat_type] += 1

                    if contrat_type == "Multisites":
                        multisites_count += 1
                    if not result:
                        non_modifiables_count += 1

                    if result:
                        self.logger.info(f"Contrat {numero_contrat} traité avec succès.")
                    else:
                        self.logger.warning(f"Échec du traitement du contrat {numero_contrat}.")
                except Exception as e:
                    self.logger.error(f"Erreur dans le thread de traitement: {e}")

        # Enregistrer les statistiques dans un fichier CSV
        csv_file_path = "resultats_traitement.csv"
        self.save_results_to_csv(
            results,
            contrat_types_count,
            multisites_count,
            non_modifiables_count,
            csv_file_path
        )

        self.logger.info("Tous les contrats ont été traités avec multi-threading.")


        self.logger.info("Tous les contrats ont été traités avec multi-threading.")

    def start(self):
        """
        Démarre le RPA avec un fichier Excel par défaut ou personnalisé.
        """
        self.main()

    def stop(self):
        """
        Arrête le processus du RPA.
        """
        self.logger.info("Arrêt du RPA Affranchigo en cours...")
        self.STOP_FLAG = True
