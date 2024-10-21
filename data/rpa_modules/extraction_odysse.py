import json
import os
import time
import csv
import threading
import sqlite3
import pandas as pd
from queue import Queue, Empty
from concurrent.futures import ThreadPoolExecutor
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium.webdriver.support.ui import Select 
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from dotenv import load_dotenv

from rpa_modules import setup_logger
from rpa_modules.data_processing import extract_contrat_numbers_to_json
# Charger les variables d'environnement

load_dotenv()
file_lock = threading.Lock()
STOP_FLAG = False

class ExtractionRPA:
    def __init__(self, pool, logger=None):
        """
        Initialise la classe avec un WebDriverPool, un logger et les identifiants.
        """
        self.pool = pool
        self.logger = logger or setup_logger("seres_case.log")
        self.url = "https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/index/acces-dc?profil=ADV"
        self.file_lock = threading.Lock()  # Lock pour les fichiers
        self.lock = threading.Lock()  # Lock pour gérer la concurrence

        # Vérification pour log le type de self.pool
        self.logger.info(f"Type de self.pool: {type(self.pool)}")


    def init_db(self, db_path="C:/Users/PHSX095/OneDrive - LA POSTE GROUPE\Documents/GUI/data/data_traitement/data_extraction.db"):
        """
        Initialise la base de données SQLite si elle n'existe pas déjà.
        """
        try:
            # Assure-toi que le chemin de la base de données est une chaîne valide
            if not isinstance(db_path, (str, bytes, os.PathLike)):
                raise TypeError(f"Le paramètre db_path doit être une chaîne de caractères, mais {type(db_path)} a été passé.")

            # Vérification si le répertoire du chemin existe, sinon le créer
            db_directory = os.path.dirname(db_path)
            if not os.path.exists(db_directory):
                os.makedirs(db_directory)
                self.logger.info(f"Répertoire créé : {db_directory}")
            else:
                self.logger.info(f"Le répertoire existe déjà : {db_directory}")

            # Se connecter à la base de données
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()

            # Créer la table si elle n'existe pas déjà
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS infos_extraites (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    numero_contrat TEXT UNIQUE,
                    client_public TEXT,
                    supp_facturation TEXT,
                    numero_EJ TEXT,
                    numero_contrat_cadre TEXT,
                    ref_bu TEXT,
                    ref_GLN TEXT,
                    numero_de_marche TEXT,
                    mode_paiement TEXT,
                    vendeur TEXT,
                    details_categorie TEXT,
                    details_type_contrat TEXT,
                    details_stype_contrat TEXT,
                    details_version_statut TEXT,
                    details_version_sstatut TEXT,
                    facturation_delai TEXT,
                    facturation_mode_paiement TEXT,
                    facturation_periodicite TEXT,
                    type_client TEXT,
                    facturation_support_facture TEXT,
                    numero_bon_commande TEXT,
                    reference_business_unit TEXT,
                    reference_site1 TEXT,
                    reference_site2 TEXT,
                    destinataire_reference_site1 TEXT,
                    destinataire_reference_site2 TEXT
                )
            ''')

            conn.commit()
            self.logger.info("Table 'infos_extraites' créée ou existante.")

            conn.close()
            self.logger.info("Base de données initialisée avec succès.")

        except sqlite3.Error as e:
            self.logger.error(f"Erreur SQLite lors de l'initialisation de la base de données: {e}")
        except Exception as e:
            self.logger.error(f"Erreur lors de l'initialisation de la base de données: {e}")
            raise


    def save_info_to_db(self, info, db_path="C:/Users/PHSX095/OneDrive - LA POSTE GROUPE/Documents/GUI/data/data_traitement/data_extraction.db"):
        """
        Sauvegarde les informations extraites dans la base de données avec mise à jour si le contrat existe déjà.
        """
        try:
            # Récupérer le chemin absolu de la base de données
            base_dir = os.path.dirname(os.path.abspath(__file__))
            db_full_path = os.path.join(base_dir, db_path)

            # Se connecter à la base de données
            conn = sqlite3.connect(db_full_path)
            cursor = conn.cursor()

            # Vérification si le numero_contrat existe déjà
            cursor.execute("SELECT COUNT(1) FROM infos_extraites WHERE numero_contrat = ?", (info['numero_contrat'],))
            exists = cursor.fetchone()[0]

            if exists:
                # Mettre à jour les informations du contrat existant
                update_keys = [f"{key} = ?" for key, value in info.items() if value is not None]
                update_values = [value for key, value in info.items() if value is not None]
                update_query = f"UPDATE infos_extraites SET {', '.join(update_keys)} WHERE numero_contrat = ?"
                
                self.logger.info(f"Requête de mise à jour générée : {update_query}")
                self.logger.info(f"Valeurs pour mise à jour : {update_values}")

                # Exécuter la requête de mise à jour
                cursor.execute(update_query, update_values + [info['numero_contrat']])
            else:
                # Insérer un nouveau contrat
                keys = [key for key, value in info.items() if value is not None]
                values = [info[key] for key in keys]
                insert_query = f"INSERT INTO infos_extraites ({', '.join(keys)}) VALUES ({', '.join(['?'] * len(values))})"

                self.logger.info(f"Requête d'insertion générée : {insert_query}")
                self.logger.info(f"Valeurs à insérer : {values}")

                # Exécuter la requête d'insertion
                cursor.execute(insert_query, values)

            # Commit les modifications
            conn.commit()
            self.logger.info("Les informations ont été insérées ou mises à jour dans la base de données avec succès.")

        except sqlite3.Error as e:
            self.logger.error(f"Erreur lors de l'insertion ou de la mise à jour dans la base de données: {e}")
        
        finally:
            if conn:
                conn.close()


    def process_json_files(self, json_path):
        """
        Lecture des numéros de contrat à partir d'un fichier JSON.
        """
        try:
            with open(json_path, 'r', encoding='utf-8') as json_file:
                contract_numbers = json.load(json_file)
                self.logger.info(f"Contenu du fichier JSON chargé : {contract_numbers}")
                return contract_numbers
        except FileNotFoundError as e:
            self.logger.error(f"Erreur lors de la lecture des numéros de contrat depuis le fichier JSON : {e}")
            return None
        except json.JSONDecodeError as e:
            self.logger.error(f"Erreur lors de l'analyse du fichier JSON : {e}")
            return None


    def handle_error_modal(self, driver):
        try:
            modal = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "body > div.bootbox.modal.fade.bootbox-alert.in > div > div"))
            )
            try:
                ok_button = modal.find_element(By.CSS_SELECTOR, "body > div.bootbox.modal.fade.bootbox-alert.in > div > div > div.modal-footer > button")
                driver.execute_script("arguments[0].click();", ok_button)
            except NoSuchElementException:
                self.logger.debug("Le bouton 'OK' n'a pas été trouvé dans la modal.")
            except Exception as e:
                self.logger.error(f"Erreur lors du clic sur le bouton 'OK' de la modal d'erreur : {e}")
        except TimeoutException:
            self.logger.debug("La modal d'erreur n'est pas apparue.")
        except Exception as e:
            self.logger.error(f"Erreur lors de la gestion de la modal d'erreur : {e}")

    def login(self, driver, wait, identifiant, mot_de_passe):
        if driver is None:
            self.logger.error("Driver non initialisé, impossible de se connecter.")
            return
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
            self.logger.debug("Déjà connecté ou le champ d'identifiant n'est pas présent.")
        except Exception as e:
            self.logger.error(f"Problème Login: {e}")

    def load_processed_contracts(self, file_path):
        if not os.path.exists(file_path):
            with open(file_path, "w") as file:
                json.dump([], file)
            self.logger.info(f"Fichier {file_path} créé car il n'existait pas.")

        try:
            with open(file_path, "r") as file:
                data = file.read().strip()
                if not data:
                    self.logger.warning("Le fichier de contrats traités est vide.")
                    return set()
                data = json.loads(data)
                if isinstance(data, list):
                    return set(data)
                else:
                    self.logger.error("Les données du fichier ne sont pas une liste.")
                    return set()
        except json.JSONDecodeError as e:
            self.logger.error(f"Erreur lors du chargement des contrats traités : {e}")
            return set()

    def save_processed_contracts(self, contrats, json_file_path):
        try:
            with open(json_file_path, "r+") as file:
                data = file.read().strip()
                existing_set = set(json.loads(data)) if data else set()
                updated_set = existing_set.union(contrats)
                file.seek(0)
                file.truncate()
                json.dump(list(updated_set), file)
        except FileNotFoundError:
            with open(json_file_path, "w") as file:
                json.dump(list(contrats), file)

    def save_non_modifiable(self, contrat_number, file_path="problemes_contrats.json"):
        data = set()
        try:
            with open(file_path, "r") as file:
                data = set(json.load(file))
        except (FileNotFoundError, json.JSONDecodeError):
            self.logger.debug("Le fichier n'existe pas ou est corrompu, il sera recréé.")

        if not isinstance(data, set):
            self.logger.error("Data n'est pas un set après le chargement depuis JSON.")
            return

        if contrat_number not in data:
            data.add(contrat_number)
            with open(file_path, "w") as file:
                json.dump(list(data), file)
            self.logger.debug(f"Contrat numéro {contrat_number} ajouté.")
        else:
            self.logger.debug(f"Contrat numéro {contrat_number} déjà présent, non ajouté.")

    def submit_contract_number(self, driver, wait, numero):
        if not driver:
            self.logger.error("Driver non initialisé, impossible de soumettre le numéro de contrat.")
            return
        self.logger.debug(f"Soumission du numéro de contrat {numero}...")
        try:
            input_element = wait.until(EC.presence_of_element_located((By.ID, "idContrat")), message="Timeout waiting for contract input field.")
            input_element.clear()
            input_element.send_keys(numero)
            input_element.send_keys(Keys.RETURN)

            submit_button = wait.until(EC.element_to_be_clickable((By.ID, "btnSubmitContrat_accesRDC")), message="Timeout waiting for submit button.")
            submit_button.click()
            self.logger.debug("Bouton de soumission cliqué avec succès.")
            self.handle_error_modal(driver)
        except TimeoutException:
            self.logger.debug("Timeout lors de la soumission du contrat.")
        except Exception as e:
            self.logger.error(f"Erreur Submit_contrat: {e}")


    def switch_to_iframe(self, driver, wait, contrat_number):
        """
        Bascule vers l'iframe pour un contrat donné.
        """
        try:
            self.handle_error_modal(driver)
            iframe_selector = "#modalRefContrat > div > div > div.modal-body > iframe"
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, iframe_selector)))
            self.logger.info(f"Changé vers l'iframe pour le contrat {contrat_number}")
        except TimeoutException:
            self.logger.error(f"L'iframe pour le contrat {contrat_number} n'est pas disponible.")
        except Exception as e:
            self.logger.error(f"{contrat_number} * Erreur lors du basculement vers l'iframe: {e}")
    
    def click_facturation_button(self, driver, wait, contrat_number):
        """
        Clique sur le bouton 'Facturation' dans l'iframe.
        """
        try:
            facturation_button_selector = "//permission/a[@href='#facturation']//div[@id='detailsFacturationButton']"
            bouton_facturation = wait.until(EC.presence_of_element_located((By.XPATH, facturation_button_selector)))
            driver.execute_script("arguments[0].click();", bouton_facturation)
            self.logger.info(f"Le bouton 'Facturation' pour le contrat {contrat_number} a été cliqué.")
        except TimeoutException:
            self.logger.error(f"Le bouton 'Facturation' pour le contrat {contrat_number} n'est pas trouvé.")
        except Exception as e:
            self.logger.error(f"{contrat_number} * Erreur lors du clic sur 'Facturation': {e}")



    def extract_details_num(self, driver, wait, contrat_number, max_retries=3):
        """
        Extrait les informations d'un contrat sur la base des éléments ID. Réessaye en cas d'échec.
        """
        retries = 0
        while retries < max_retries:
            try:
                # Attendre que les éléments soient disponibles et extraire leurs valeurs
                detailsCategorieV = wait.until(EC.presence_of_element_located((By.ID, "detailsCategorieV"))).text.strip()
                detailsTypeContratV = wait.until(EC.presence_of_element_located((By.ID, "detailsTypeContratV"))).text.strip()
                detailsSTypeContratV = wait.until(EC.presence_of_element_located((By.ID, "detailsSTypeContratV"))).text.strip()
                detailsVersionStatutV = wait.until(EC.presence_of_element_located((By.ID, "detailsVersionStatutV"))).text.strip()
                detailsVersionSStatutV = wait.until(EC.presence_of_element_located((By.ID, "detailsVersionSStatutV"))).text.strip()

                # Log des résultats
                self.logger.info(f"{contrat_number} * Détails extraits: {detailsCategorieV}, {detailsTypeContratV}, {detailsSTypeContratV}, {detailsVersionStatutV}, {detailsVersionSStatutV}")

                # Sauvegarder dans la base de données
                self.save_info_to_db({
                    "numero_contrat": contrat_number,
                    "details_categorie": detailsCategorieV,
                    "details_type_contrat": detailsTypeContratV,
                    "details_stype_contrat": detailsSTypeContratV,
                    "details_version_statut": detailsVersionStatutV,
                    "details_version_sstatut": detailsVersionSStatutV
                })
                
                break  # Si tout est bien extrait, sortir de la boucle
                
            except Exception as e:
                retries += 1
                if retries >= max_retries:
                    self.logger.error(f"Erreur critique lors de l'extraction des détails du contrat {contrat_number} après {max_retries} tentatives: {e}")
                    raise  # Relancer l'erreur après le nombre maximal de retries
                else:
                    self.logger.warning(f"Tentative {retries}/{max_retries} échouée pour l'extraction des détails du contrat {contrat_number}: {e}")
                    self.save_non_modifiable(contrat_number)

    
    def extract_facturation_element(self, driver, wait, contrat_number, max_retries=3):
        """
        Extrait les éléments de facturation, incluant les sélecteurs et la checkbox, avec gestion stricte et retries.
        """
        retries = 0
        while retries < max_retries:
            try:
                # Partie 1 : Extraction des sélecteurs
                facturation_delai = Select(wait.until(EC.presence_of_element_located((By.ID, "facturationPaiemenDelaiI")))).first_selected_option.text.strip()
                facturation_mode_paiement = Select(wait.until(EC.presence_of_element_located((By.ID, "facturationPaiementModePaiementI")))).first_selected_option.text.strip()
                facturation_periodicite = Select(wait.until(EC.presence_of_element_located((By.ID, "facturationPaiementPeriodiciteI")))).first_selected_option.text.strip()

                self.logger.info(f"{contrat_number} * Facturation Délai: {facturation_delai}, Mode Paiement: {facturation_mode_paiement}, Périodicité: {facturation_periodicite}")

                # Partie 2 : Extraction de la valeur de la checkbox
                type_client_checkbox = wait.until(EC.presence_of_element_located((By.ID, "facturationPaiementTypeClientI"))).is_selected()
                type_client_value = "True" if type_client_checkbox else "False"
                
                self.logger.info(f"{contrat_number} * Checkbox Type Client: {type_client_value}")

                # Partie 3 : Extraction de l'élément facturationPaiementSupportFactureI (option sélectionnée uniquement)
                facturation_support_facture = Select(wait.until(EC.presence_of_element_located((By.ID, "facturationPaiementSupportFactureI")))).first_selected_option.text.strip()

                self.logger.info(f"{contrat_number} * Support Facture sélectionné: {facturation_support_facture}")
                
                # Sauvegarder dans la base de données
                self.save_info_to_db({
                    "numero_contrat": contrat_number,
                    "facturation_delai": facturation_delai,
                    "facturation_mode_paiement": facturation_mode_paiement,  # Vérifiez bien cette clé
                    "facturation_periodicite": facturation_periodicite,
                    "type_client": type_client_value,
                    "facturation_support_facture": facturation_support_facture
                })

                self.logger.info(f"{contrat_number} * Facturation: {facturation_delai}, {facturation_mode_paiement}, {facturation_periodicite}, {type_client_value}, {facturation_support_facture}")

                break  # Si tout est bien extrait, sortir de la boucle

            except Exception as e:
                retries += 1
                if retries >= max_retries:
                    self.logger.error(f"Erreur critique lors de l'extraction des éléments de facturation pour le contrat {contrat_number} après {max_retries} tentatives: {e}")
                    raise  # Relancer l'erreur après le nombre maximal de retries
                else:
                    self.logger.warning(f"Tentative {retries}/{max_retries} échouée pour l'extraction des éléments de facturation du contrat {contrat_number}: {e}")
                    self.save_non_modifiable(contrat_number)
                    


    def click_reference_client_tab(self, driver, wait):
        """
        Fonction pour cliquer sur l'onglet 'Référence Client'.
        """
        try:
            # Utiliser WebDriverWait pour attendre que l'onglet soit cliquable
            reference_client_tab = wait.until(EC.element_to_be_clickable((By.ID, "ReferenceClientTitle")))
            reference_client_tab.click()
            self.logger.info("Onglet 'Référence Client' cliqué avec succès.")
            
            # Attendre un peu après le clic pour s'assurer que le contenu est chargé
            time.sleep(2)
            
        except Exception as e:
            self.logger.error(f"Erreur lors du clic sur l'onglet 'Référence Client': {e}")

    def extract_element_reference_client(self, driver, wait, contrat_number, max_retries=3):
        """
        Extrait les éléments de l'onglet 'Référence Client' et les sauvegarde dans la base de données avec retries.
        """
        retries = 0
        while retries < max_retries:
            try:
                # Fonction pour extraire le texte d'un élément ou retourner None si l'élément est absent
                def get_element_text_by_id(element_id):
                    try:
                        element = wait.until(EC.presence_of_element_located((By.ID, element_id)))
                        text = element.text.strip()
                        return text if text else None
                    except Exception as e:
                        self.logger.warning(f"Élément {element_id} non trouvé ou vide.")
                        return None

                # Extraction des éléments
                numero_bon_commande = get_element_text_by_id("facturationClientNumeroBonCommandeI")
                reference_business_unit = get_element_text_by_id("facturationClientReferenceBusinessUnitI")
                reference_site1 = get_element_text_by_id("facturationClientReferenceSite1I")
                reference_site2 = get_element_text_by_id("facturationClientReferenceSite2I")
                destinataire_reference_site1 = get_element_text_by_id("facturationClientDestinataireReferenceSite1I")
                destinataire_reference_site2 = get_element_text_by_id("facturationClientDestinataireReferenceSite2I")
                numero_marche = get_element_text_by_id("facturationClientNumeroMarcheI")

                # Sauvegarder dans la base de données
                self.save_info_to_db({
                    "numero_contrat": contrat_number,
                    "numero_bon_commande": numero_bon_commande,
                    "reference_business_unit": reference_business_unit,
                    "reference_site1": reference_site1,
                    "reference_site2": reference_site2,
                    "destinataire_reference_site1": destinataire_reference_site1,
                    "destinataire_reference_site2": destinataire_reference_site2,
                    "numero_de_marche": numero_marche
                })

                self.logger.info(f"{contrat_number} * Référence Client: Bon de commande: {numero_bon_commande}, BU: {reference_business_unit}, "
                                f"Site 1: {reference_site1}, Site 2: {reference_site2}, Destinataire Site 1: {destinataire_reference_site1}, "
                                f"Destinataire Site 2: {destinataire_reference_site2}, Numéro de Marché: {numero_marche}")

                break  # Si tout est bien extrait, sortir de la boucle

            except Exception as e:
                retries += 1
                if retries >= max_retries:
                    self.logger.error(f"Erreur critique lors de l'extraction des éléments de 'Référence Client' pour le contrat {contrat_number} après {max_retries} tentatives: {e}")
                    raise  # Relancer l'erreur après le nombre maximal de retries
                else:
                    self.logger.warning(f"Tentative {retries}/{max_retries} échouée pour l'extraction des éléments de 'Référence Client' du contrat {contrat_number}: {e}")
                    self.save_non_modifiable(contrat_number)


    def process_contract(self, numero_contrat, identifiant, mot_de_passe, processed_contracts, retry_count=3):
        retry_attempts = 0
        while retry_attempts < retry_count:
            try:
                driver = self.pool.get_driver(self.url)  
                self.logger.info(f"Début du traitement du contrat {numero_contrat}")
                start_time = time.time()

                # Vérification si le contrat a déjà été traité
                with self.lock:
                    if numero_contrat in processed_contracts:
                        self.logger.info(f"Le contrat {numero_contrat} a déjà été traité.")
                        return (numero_contrat, "Déjà traité", 0)

                # 1. Connexion
                self.logger.info(f"Tentative de connexion pour le contrat {numero_contrat}")
                self.login(driver, WebDriverWait(driver, 20), identifiant, mot_de_passe)

                # 2. Soumettre le numéro de contrat
                self.logger.info(f"Soumission du numéro de contrat {numero_contrat}")
                self.submit_contract_number(driver, WebDriverWait(driver, 20), numero_contrat)

                # 3. Extraction des détails du contrat
                self.logger.info(f"Extraction des détails pour le contrat {numero_contrat}")
                self.switch_to_iframe(driver, WebDriverWait(driver, 20), numero_contrat)
                self.extract_details_num(driver, WebDriverWait(driver, 20), numero_contrat)

                # 4. Extraction des éléments de facturation
                self.logger.info(f"Extraction des éléments de facturation pour le contrat {numero_contrat}")
                self.click_facturation_button(driver, WebDriverWait(driver, 20), numero_contrat)
                self.extract_facturation_element(driver, WebDriverWait(driver, 20), numero_contrat)

                # 5. Extraction des références client
                self.logger.info(f"Extraction des références client pour le contrat {numero_contrat}")
                self.click_reference_client_tab(driver, WebDriverWait(driver, 20))
                self.extract_element_reference_client(driver, WebDriverWait(driver, 20), numero_contrat)

                # Sauvegarde de l'état du contrat traité
                self.logger.info(f"Le contrat {numero_contrat} a été traité avec succès.")
                with self.lock:
                    processed_contracts.add(numero_contrat)
                    self.save_processed_contracts(processed_contracts, "numeros_contrat_traites.json")

                # Calcul de la durée du traitement
                end_time = time.time()
                duration = end_time - start_time
                return (numero_contrat, "Mise à jour réussie", duration)

            except Exception as e:
                retry_attempts += 1
                self.logger.error(f"Erreur pour le contrat {numero_contrat} : {e}")
                if retry_attempts >= retry_count:
                    self.save_non_modifiable(numero_contrat)
                    return (numero_contrat, "Échec après plusieurs tentatives", 0)
            finally:
                self.pool.return_driver(driver)  # Retourne le WebDriver au pool


    def worker(self, queue, progress_callback, total_contracts, identifiant, mot_de_passe):
        processed_contracts = self.load_processed_contracts("numeros_contrat_traites.json")
        results = []
        processed_count = 0

        while not self.STOP_FLAG:
            try:
                numero_contrat = queue.get(timeout=1)
                if numero_contrat is None:
                    break

                result = self.process_contract(numero_contrat, identifiant, mot_de_passe, processed_contracts)
                if result:
                    results.append(result)
                    processed_count += 1
                    if progress_callback:
                        progress_callback((processed_count / total_contracts) * 100)

                queue.task_done()

            except Exception as e:
                self.logger.error(f"Erreur dans le worker: {e}")

        return results

    def process_batches(self, contract_numbers, identifiant, mot_de_passe, progress_callback=None, batch_size=500):
        # Convertir les valeurs du dictionnaire en liste
        contract_values = list(contract_numbers.values())  # Extraire les valeurs réelles
        total_contracts = len(contract_values)
        all_results = []
        num_batches = (total_contracts // batch_size) + (1 if total_contracts % batch_size > 0 else 0)

        for i in range(num_batches):
            # Utiliser le slicing sur les valeurs du dictionnaire
            batch_contracts = contract_values[i * batch_size:(i + 1) * batch_size]
            self.logger.info(f"Traitement du batch {i + 1}/{num_batches} ({len(batch_contracts)} contrats)")

            queue = Queue()
            for numero_contrat in batch_contracts:
                queue.put(numero_contrat)

            with ThreadPoolExecutor(max_workers=self.pool.max_size) as executor:
                futures = [executor.submit(self.worker, queue, progress_callback, total_contracts, identifiant, mot_de_passe) for _ in range(self.pool.max_size)]
                for future in futures:
                    try:
                        results = future.result()
                        if results:
                            all_results.extend(results)
                    except Exception as e:
                        self.logger.error(f"Erreur lors du traitement du contrat: {e}")

            queue.join()
            self.pool.close_all()  # Fermer les WebDrivers après chaque batch pour libérer les ressources

        return all_results

    
    def process_contract_in_parallel(self, contract_numbers, identifiant, mot_de_passe, max_workers=5):
        """
        Fonction pour traiter plusieurs contrats en parallèle en utilisant ThreadPoolExecutor.
        :param contract_numbers: Dictionnaire des numéros de contrat à traiter.
        :param identifiant: Identifiant pour se connecter à la plateforme.
        :param mot_de_passe: Mot de passe pour se connecter à la plateforme.
        :param max_workers: Nombre maximum de threads (WebDrivers) à utiliser en parallèle.
        """
        # Liste pour stocker les résultats des contrats traités
        all_results = []

        # Pool de threads pour traiter les contrats en parallèle
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(self.process_contract, numero_contrat, identifiant, mot_de_passe, self.load_processed_contracts("numeros_contrat_traites.json")): numero_contrat
                for numero_contrat in contract_numbers.values()  # Utilisez les valeurs réelles des numéros de contrat
            }

            # Traiter les contrats au fur et à mesure que les threads terminent
            for future in as_completed(futures):
                numero_contrat = futures[future]
                try:
                    result = future.result()  # Obtenir le résultat du thread terminé
                    self.logger.info(f"Contrat {numero_contrat} traité avec succès : {result}")
                    all_results.append(result)
                except Exception as exc:
                    self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat}: {exc}")

        # Retourner tous les résultats des contrats traités
        return all_results


    def save_results_to_csv(self, results, csv_file_path):
        if not results:
            self.logger.error("Aucun résultat à enregistrer.")
            return

        headers = ['Numéro de contrat', 'Résultat', 'Durée (secondes)']
        with open(csv_file_path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(headers)
            for result in results:
                writer.writerow(result)
            self.logger.info(f"Résultats enregistrés dans {csv_file_path}")


    def main(self, excel_path=None, progress_callback=None):
        """
        Méthode principale pour le traitement du RPA Extraction avec multi-traitement.
        """
        self.logger.info("Démarrage du script d'extraction...")

        base_dir = os.path.dirname(os.path.abspath(__file__))

        # Utiliser un chemin par défaut si aucun chemin n'est spécifié
        if excel_path is None:
            excel_path = os.path.join(base_dir, '../data_traitement', 'Affranchigo_Demat_helene.xlsx')

        if not os.path.exists(excel_path):
            self.logger.error(f"Le fichier Excel spécifié est introuvable : {excel_path}")
            raise FileNotFoundError(f"Le fichier {excel_path} est introuvable.")

        # Initialiser la base de données
        self.init_db()

        # Chemin pour enregistrer le fichier JSON
        json_dir = os.path.join(base_dir, '../data_traitement')
        json_path = os.path.join(json_dir, 'numeros_contrat_extraction_cf.json')

        # Extraire les numéros de contrat à partir du fichier Excel et les sauvegarder dans un fichier JSON
        self.logger.info(f"Extraction des numéros de contrat depuis {excel_path}")
        extract_contrat_numbers_to_json(excel_path, json_path)

        # Charger les numéros de contrat depuis le fichier JSON
        contract_numbers = self.process_json_files(json_path)
        if contract_numbers is None:
            self.logger.error("Aucun numéro de contrat chargé depuis le fichier JSON.")
            return

        # Charger les identifiants
        identifiant = os.getenv("IDENTIFIANT")
        mot_de_passe = os.getenv("MOT_DE_PASSE")

        # Multi-traitement des contrats en parallèle
        results = self.process_contract_in_parallel(contract_numbers, identifiant, mot_de_passe, max_workers=5)

        # Sauvegarder les résultats dans un fichier CSV
        self.save_results_to_csv(results, "results.csv")

        # Fermer tous les WebDrivers du pool
        self.pool.close_all()



    def start(self, excel_path=None):
        """
        Démarre le RPA avec un fichier Excel par défaut ou personnalisé.
        """
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        if excel_path is None:
            excel_path = os.path.join(base_dir, 'data_traitement', 'Affranchigo_Demat_helene.xlsx')

        self.logger.info(f"Lancement du RPA Extraction avec le fichier {excel_path}")
        self.STOP_FLAG = False
        self.main(excel_path)

