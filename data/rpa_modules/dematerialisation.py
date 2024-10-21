import os
import json
import time
import csv
from dotenv import load_dotenv
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from queue import Queue 
import queue

from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from rpa_modules.data_processing import extract_contrat_numbers_to_json


from rpa_modules import setup_logger

load_dotenv()

class CasDematerialisationRPA:
    def __init__(self, pool, logger=None):
        """
        Initialise la classe avec un pool de WebDrivers et un logger.
        :param pool: Instance de WebDriverPool.
        :param logger: Logger pour les logs. Un logger par défaut est créé si non fourni.
        """
        self.pool = pool
        self.logger = logger or setup_logger("demat_case.log")
        self.STOP_FLAG = False
        self.processed_count = 0
        self.error_count = 0
        self.start_time = None


    def save_metrics_to_csv(self, results, total_error_count, file_path="metrics.csv"):
        # Sauvegarder les métriques dans un fichier CSV
        with open(file_path, mode='a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Date", "Numéro Contrat", "Résultat", "Temps de traitement (s)"])

            for result in results:
                numero_contrat, status, contract_time = result
                writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), numero_contrat, status, contract_time])

            # Ajouter le nombre total d'erreurs à la fin
            writer.writerow(["Total Errors", total_error_count])
        
        self.logger.info(f"Métriques sauvegardées dans {file_path}")


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
            self.logger.debug("Dejà connecté ou le champ d'identifiant n'est pas présent.")
        except Exception as e:
            self.logger.error(f"Problème Login : {e}")

    def process_json_files(self, file_path):
        self.logger.info("Traitement du fichier JSON pour les contrats...")
        if os.path.exists(file_path):
            with open(file_path, "r") as file:
                data = json.load(file)
                return list(data.values())
        return []

    def submit_contract_number(self, driver, wait, numero):
        self.logger.info(f"Soumission du numéro de contrat {numero}...")
        try:
            input_element = wait.until(EC.presence_of_element_located((By.ID, "idContrat")))
            input_element.clear()
            input_element.send_keys(numero)
            input_element.send_keys(Keys.RETURN)

            submit_button = wait.until(EC.element_to_be_clickable((By.ID, "btnSubmitContrat_accesRDC")))
            submit_button.click()

            wait.until(EC.visibility_of_element_located((By.ID, "modalRefContrat")))
        except TimeoutException:
            self.logger.error("Problème avec la soumission du numéro de contrat.")
        except Exception as e:
            self.logger.error(f"Erreur soumission contrat : {e}")

    def save_processed_contracts(self, contrats, file_path="numeros_contrat_traites_extraction.json"):
        try:
            if os.path.exists(file_path):
                with open(file_path, "r+") as file:
                    existing_data = json.load(file)
                    updated_data = list(set(existing_data + contrats))
                    file.seek(0)
                    json.dump(updated_data, file)
            else:
                with open(file_path, "w") as file:
                    json.dump(contrats, file)
        except Exception as e:
            self.logger.error(f"Erreur sauvegarde contrats : {e}")
    
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

    def save_non_modifiable_contract(self, contrat_number, file_path="cas_non_modifiable.json"):
        try:
            if os.path.exists(file_path):
                with open(file_path, "r+") as file:
                    data = json.load(file)
                    if contrat_number not in data:
                        data.append(contrat_number)
                        file.seek(0)
                        json.dump(data, file)
            else:
                with open(file_path, "w") as file:
                    json.dump([contrat_number], file)
        except Exception as e:
            self.logger.error(f"Erreur sauvegarde contrat non modifiable : {e}")
    
    def write_version_comment(self, driver, wait, numeros_contrat):
        try:
            # Attendre que l'élément de commentaire soit présent et visible
            comment_element = wait.until(EC.visibility_of_element_located((By.ID, "versionComment")))
            
            # Effacer tout texte existant dans l'élément (optionnel, si nécessaire)
            comment_element.clear()
            
            # Écrire "Support PDF signé" dans l'élément
            comment_element.send_keys(f"{numeros_contrat}Support PDF signé")
            
            self.logger.info("Texte 'Support PDF signé' écrit avec succès dans le champ de commentaire.")
        
        except Exception as e:
            self.logger.error(f"Erreur lors de l'écriture dans le champ 'versionComment': {e}")
        
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
    
    def click_submit_button_and_confirm(self, driver, wait, numero_contrat):
        try:
            # Attendre que le premier bouton soit cliquable et cliquer dessus
            submit_button = wait.until(EC.element_to_be_clickable((By.ID, "detailsButtonValidationFacturation")))
            submit_button.click()
            self.logger.info(f"{numero_contrat}*Premier bouton 'Validation Facturation' cliqué avec succès.")
            
            # Attendre que le deuxième bouton Swal apparaisse et soit cliquable
            second_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div.swal2-container.swal2-center.swal2-fade.swal2-shown > div > div.swal2-buttonswrapper > button.swal2-confirm.swal2-styled")))
            second_button.click()
            self.logger.info(f"{numero_contrat}*Deuxième bouton 'Swal2 Confirm' cliqué avec succès.")
        
        except Exception as e:
            self.logger.error(f"{numero_contrat}*Erreur lors du clic sur les boutons de validation : {e}")
            self.save_non_modifiable_contract(numero_contrat)
        
        finally:
            # Redirection vers l'URL de base après le clic sur le deuxième bouton, que ça réussisse ou non
            driver.get("https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/index/acces-dc?profil=ADV")
            self.logger.info("Redirection vers l'URL de base après le clic.")


    def select_pdf_signe(self, driver, wait, numero_contrat):
        try:
            # Attendre que l'élément <select> soit visible et interactif
            select_element = wait.until(EC.element_to_be_clickable((By.ID, "facturationPaiementSupportFactureI")))
            
            # Créer un objet Select à partir de l'élément
            select = Select(select_element)
            
            # Sélectionner l'option par index (index 3 correspondant à "PDF Signé")
            select.select_by_index(3)
            
            self.logger.info(f"{numero_contrat}*Option 'PDF Signé' sélectionnée avec succès.")
        
        except Exception as e:
            self.logger.error(f"{numero_contrat}*Erreur lors de la sélection de l'option 'PDF Signé': {e}")
            self.save_non_modifiable_contract(numero_contrat)
    
    def process_batches(self, contract_numbers, batch_size, identifiant, mot_de_passe):
        total_contracts = len(contract_numbers)
        num_batches = (total_contracts + batch_size - 1) // batch_size
        results = []
        total_error_count = 0  # Initialiser un compteur global d'erreurs

        for batch_num in range(num_batches):
            start_idx = batch_num * batch_size
            end_idx = min(start_idx + batch_size, total_contracts)
            batch_contracts = contract_numbers[start_idx:end_idx]

            self.logger.debug(f"Traitement du lot {batch_num + 1}/{num_batches} : contrats {start_idx + 1} à {end_idx}")

            queue = Queue()
            for numero in batch_contracts:
                if numero is None:
                    continue
                self.logger.debug(f"Ajout du contrat {numero} à la queue.")
                queue.put(numero)

            with ThreadPoolExecutor(max_workers=self.pool.max_size) as executor:
                futures = [
                    executor.submit(self.worker, queue, identifiant, mot_de_passe)
                    for _ in range(self.pool.max_size)
                ]

                for future in futures:
                    try:
                        batch_results, batch_error_count = future.result()
                        results.extend(batch_results)
                        total_error_count += batch_error_count  # Incrémenter les erreurs du lot
                    except Exception as e:
                        self.logger.error(f"Erreur lors du traitement d'un contrat dans le lot {batch_num + 1}: {e}")

            queue.join()
            self.logger.debug(f"Lot {batch_num + 1}/{num_batches} traité avec succès.")

        # Retourner les résultats avec le nombre total d'erreurs
        return results, total_error_count


    
    def worker(self, queue, identifiant, mot_de_passe):
        results = []
        error_count = 0  # Initialiser un compteur d'erreurs
        try:
            driver = self.pool.get_driver("https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/index/acces-dc?profil=ADV")
            while not self.STOP_FLAG:
                try:
                    numero_contrat = queue.get(timeout=1)
                    if numero_contrat is None:
                        break

                    # Commencer le chronométrage pour ce contrat
                    start_time = time.time()

                    # Traiter le contrat avec le WebDriver récupéré
                    result = self.process_contract(driver, numero_contrat, identifiant, mot_de_passe)

                    # Fin du chronométrage et calcul du temps de traitement
                    contract_time = time.time() - start_time

                    # Ajouter le résultat avec le temps pris et l'état (succès ou erreur)
                    if result[1] == 'Erreur':
                        error_count += 1  # Incrémenter le compteur d'erreurs
                    results.append((numero_contrat, result[1], contract_time))  # (numéro de contrat, succès/erreur, temps)
                    queue.task_done()

                    if self.STOP_FLAG:
                        break
                except queue.Empty:
                    break
                except Exception as e:
                    self.logger.error(f"Erreur dans le worker : {e}")
                    error_count += 1
        finally:
            self.pool.return_driver(driver)  # Retourner le WebDriver au pool après utilisation
            self.logger.debug("Worker terminé.")
        
        return results, error_count


    def process_contract(self, driver, numero_contrat, identifiant, mot_de_passe):
        try:
            self.logger.info(f"Tentative de login pour le contrat {numero_contrat}")
            wait = WebDriverWait(driver, 10)
            self.login(driver, wait, identifiant, mot_de_passe)
            
            self.logger.info(f"Soumission du numéro de contrat {numero_contrat}")
            self.submit_contract_number(driver, wait, numero_contrat)
 
            self.logger.info(f"Passage à l'iframe pour le contrat {numero_contrat}")
            self.switch_to_iframe(driver, wait, numero_contrat)
           
            self.logger.info(f"Cliquer sur Facturation pour le contrat {numero_contrat}")
            self.click_facturation_button(driver, wait, numero_contrat)
            
            self.logger.info(f"Sélectionner PDF Signé pour le contrat {numero_contrat}")
            self.select_pdf_signe(driver, wait, numero_contrat)
            time.sleep(3)
            self.logger.info(f"Écriture du commentaire version pour le contrat {numero_contrat}")
            self.write_version_comment(driver, wait, numero_contrat)
           
            self.logger.info(f"Cliquer sur le bouton de soumission pour le contrat {numero_contrat}")
            self.click_submit_button_and_confirm(driver, wait, numero_contrat)
          
            self.logger.info(f"{numero_contrat} * Traitement terminé.")
            self.save_processed_contracts([numero_contrat])

            return (numero_contrat, True, "Succès")
        except Exception as e:
            self.logger.critical(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")
            self.save_non_modifiable_contract(numero_contrat)
            return (numero_contrat, False, "Erreur")
        finally:
            driver.get("https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/index/acces-dc?profil=ADV")


    def main(self, excel_path="C:/Users/PHSX095/OneDrive - LA POSTE GROUPE/Documents/GUI/data/data_traitement/Affranchigo_Demat_helene.xlsx"):
        self.logger.info("Démarrage du script pour la dématérialisation...")

        # Initialiser la variable start_time
        self.start_time = time.time()

        identifiant = os.getenv("IDENTIFIANT")
        mot_de_passe = os.getenv("MOT_DE_PASSE")

        # Extraire les numéros de contrats depuis l'Excel et les sauvegarder en JSON
        json_file_path = "cas_privée_a_traiter.json"
        extract_contrat_numbers_to_json(excel_path, json_file_path)

        # Charger les numéros de contrats à traiter depuis le fichier JSON
        numeros_contrat = self.process_json_files(json_file_path)

        # Traiter les contrats en lots
        results, total_error_count = self.process_batches(numeros_contrat, 10, identifiant, mot_de_passe)

        # Sauvegarder les métriques après traitement
        self.save_metrics_to_csv(results, total_error_count)

        return results



    def start(self, excel_path="C:/Users/PHSX095/OneDrive - LA POSTE GROUPE/Documents/GUI/data/data_traitement/Affranchigo_Demat_helene.xlsx"):
        self.logger.info(f"Lancement du RPA Dématerialisation avec le fichier {excel_path}")
        self.STOP_FLAG = False
        self.main(excel_path)

    def stop(self):
        self.logger.info("Arrêt du RPA Dématerialisation en cours...")
        self.STOP_FLAG = True