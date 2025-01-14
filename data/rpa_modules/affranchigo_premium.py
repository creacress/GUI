from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webelement import WebElement
import json
import pandas as pd
import time

from rpa_modules import setup_logger

class AffranchigoPremiumCase:
    def __init__(self, driver, pool, logger=None):
        """
        Initialise la classe avec un WebDriver, un pool, et un logger.
        :param driver: L'instance de WebDriver à utiliser.
        :param pool: Le pool de WebDrivers partagé.
        :param logger: Logger pour les logs. Un logger par défaut est créé si non fourni.
        """
        self.driver = driver  # Initialisation du WebDriver
        self.pool = pool
        self.logger = logger or setup_logger("affranchigo_lib_case.log")

    def get_selector_value(self, selector):
        """
        Fonction utilitaire pour obtenir la valeur d'un input ou d'un sélecteur.
        
        :param selector: L'élément dont la valeur doit être extraite.
        :return: La valeur extraite de l'élément.
        """
        if selector.tag_name == 'select':
            # Pour un sélecteur avec options, on retourne la valeur de l'option sélectionnée
            return selector.find_element_by_css_selector('option:checked').get_attribute('title')
        else:
            # Pour un input, on retourne la valeur de l'input
            return selector.get_attribute('value')
    
    def initialize_selectors(self, numero_contrat, max_retries=3): 
        selectors = {}
        retries = 0

        while retries < max_retries:
            try:
                self.logger.debug("Initialisation des sélecteurs.")
                wait = WebDriverWait(self.driver, 20)  # Augmentation du délai d'attente

                selectors["input_first_regate"] = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#g0_p10858\\|0_r72238_c72239\\[0\\]")))
                self.logger.debug("Sélecteur input_first_regate initialisé.")

                selectors["select_first_etablissement"] = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#\\[g0_p10858\\|0_r72238\\[0\\]\\] > div > critere-form:nth-child(3) > div.form-group.critere_psc > input-etb-prest > div > select")))
                self.logger.debug("Sélecteur select_first_etablissement initialisé.")
                
                try:
                    selectors["select_first_role"] = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#g0_p10858\\|0_r72238_c72243\\[0\\]_v1"))) 
                    self.logger.debug("Sélecteur select_first_role  (1er) initialisé.")
                except:
                    selectors["select_first_role"] = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#g0_p10858\\|0_r72238_c72243\\[0\\]_v2"))) 
                    self.logger.debug("Sélecteur select_first_role (2 ème) initialisé.")


                selectors["input_second_regate"] = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#g0_p10858\\|0_r72238_c72239\\[1\\]")))
                self.logger.debug("Sélecteur input_second_regate initialisé.")

                selectors["select_second_etablissement"] = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#\\[\\[g0_p10858\\|0_r72238\\[0\\]\\]\\] > div > critere-form:nth-child(3) > div.form-group.critere_psc > input-etb-prest > div > select")))
                self.logger.debug("Sélecteur select_second_etablissement initialisé.")

                try:

                    selectors["select_second_role"] = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#g0_p10858\\|0_r72238_c72243\\[1\\]_v2")))
                    self.logger.debug("Sélecteur select_second_role initialisé.")
                except:
                    selectors["select_second_role"] = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#g0_p10858\\|0_r72238_c72243\\[1\\]_v1")))
                    self.logger.debug("Sélecteur select_second_role initialisé.")

                # Si tous les sélecteurs sont initialisés, on retourne le dictionnaire
                return selectors

            except Exception as e:
                retries += 1
                self.logger.warning(f"{numero_contrat} * Tentative {retries}/{max_retries} échouée en raison d'un timeout. Réessayer...")
                if retries == max_retries:
                    self.logger.error(f"{numero_contrat} * Timeout lors de l'initialisation des sélecteurs après {max_retries} tentatives: {e}")
                return None

        # Si toutes les tentatives échouent, retourner None
        return None


    def select_time_in_selectors(self, numero_contrat):
        select_time_selectors = [
            "#\\[g0_p10858\\|0_r72238\\[0\\]\\] > div > critere-form:nth-child(7) > div.form-group.critere_psc > input-component > div > select",
            "#\\[g0_p10858\\|0_r72238\\[0\\]\\] > div > critere-form:nth-child(5) > div.form-group.critere_psc > input-component > div > select",
            "#\\[\\[g0_p10858\\|0_r72238\\[0\\]\\]\\] > div > critere-form:nth-child(7) > div.form-group.critere_psc > input-component > div > select",
            "#\\[\\[g0_p10858\\|0_r72238\\[0\\]\\]\\] > div > critere-form:nth-child(5) > div.form-group.critere_psc > input-component > div > select"
        ]
        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, ",".join(select_time_selectors))
                )
            )
        
            for select_time_selector in select_time_selectors:
                for element in self.driver.find_elements(By.CSS_SELECTOR, select_time_selector):
                    try:
                        select_element = Select(element)
                        current_value = select_element.first_selected_option.get_attribute('value')

                        if current_value in ['', 'null']:
    
                            select_element.select_by_index(16)
                        else:
                            self.logger.info(f"{numero_contrat} * L'heure est déjà sélectionnée: '{current_value}'.")
                    except Exception as e:
                        self.logger.critical(f"{numero_contrat} * Erreur lors de la sélection de l'heure: {e}")
        except Exception as e:
            self.logger.exception(f"{numero_contrat} * Aucun sélecteur de temps trouvé. {e}")
            pass
    
    def save_non_modifiable_contract(self, numero_contrat):
        """Enregistre le numéro de contrat dans un fichier JSON en évitant les doublons."""
        file_path = "re_traitement_lib.json"
        data = set()  # Utiliser un ensemble pour éviter les doublons

        try:
            self.logger.debug(f"Lecture du fichier existant : {file_path}")
            with open(file_path, "r") as file:
                existing_data = json.load(file)
                data.update(existing_data)  # Mise à jour de l'ensemble avec les données existantes
        except FileNotFoundError:
            self.logger.debug("Le fichier n'existe pas, il sera créé.")
        except json.JSONDecodeError:
            self.logger.error("Erreur de décodage JSON, le fichier est peut-être corrompu.")

        if numero_contrat not in data:
            data.add(numero_contrat)
            with open(file_path, "w") as file:
                json.dump(list(data), file)  # Convertir l'ensemble en liste pour la sérialisation
            self.logger.debug(f"Contrat numéro {numero_contrat} ajouté.")
        else:
            self.logger.debug(f"Contrat numéro {numero_contrat} déjà présent, non ajouté.")


    def submit_liberte(self, numero_contrat):
        button_selectors = [
            "#odcFormCPV > button",
            "#odcFormCPR > button",
            "#odcFormCPT > button"
        ]

        button_found = False
        for selector in button_selectors:
            try:
                # Localiser le bouton
                submit_button = WebDriverWait(self.driver, 10).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, selector))
                )

                # Scroller jusqu'au bouton
                self.driver.execute_script("arguments[0].scrollIntoView(true);", submit_button)
                time.sleep(1)  # Pause optionnelle pour le scroll

                # Clic via JavaScript
                self.driver.execute_script("arguments[0].click();", submit_button)

                self.logger.info(f"{numero_contrat} * Clic sur le bouton de soumission avec succès.")
                button_found = True
                break
            except TimeoutException:
                self.logger.debug(f"Le bouton avec le sélecteur {selector} n'a pas été trouvé.")
            except Exception as e:
                self.logger.error(f"Erreur lors de la tentative de clic sur le bouton {selector}: {e}")
        
        if not button_found:
            self.logger.error(f"{numero_contrat} * Aucun bouton de soumission n'a été trouvé.")
            return

        try:
            # Attente du changement d'URL après la soumission
            WebDriverWait(self.driver, 10).until(
                EC.url_changes(
                    "https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/configurateur/put-contract"
                )
            )
            self.logger.info(f"{numero_contrat} * Formulaire soumis avec succès.")
            time.sleep(3)

            # Navigation vers l'URL de départ
            try:
                url_de_depart = "https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/index/acces-dc"
                self.driver.get(url_de_depart)
                self.logger.debug("Retour à l'URL de départ réussi.")
            except Exception as e:
                self.logger.critical(f"Erreur lors de la navigation vers l'URL de départ : {e}")
        except TimeoutException:
            self.logger.error("Le bouton de soumission n'a pas été trouvé dans les temps.")
        except Exception as e:
            self.save_non_modifiable_contract(numero_contrat)
            self.logger.critical(f"Erreur lors de la soumission du formulaire : {e}")



    def check_div_alert(self):
        try:
            return WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#g0_p159\\|0_r486\\[0\\] > div > critere-form:nth-child(3) > error-component > div.alert.alert-danger"))
            )
        except TimeoutException:
            return None

    def update_input(self, element, new_value, numero_contrat):
        self.logger.debug(f"Tentative de mise à jour de l'input avec la nouvelle valeur: {new_value}")
        if element is None:
            self.logger.error("L'élément fourni à update_input est None.")
            return
        if new_value is None:
            self.logger.error("La nouvelle valeur pour l'input est None.")
            return
        element.clear()
        element.send_keys(new_value)
        element.send_keys(Keys.TAB)
        self.logger.debug(f"{numero_contrat} * Nouveau code régate attribué {new_value}")

    def update_select_element(self, driver, select_element, numero_contrat):
        self.logger.debug(f"Mise à jour de l'élément select: {select_element}")

        try:
            if isinstance(select_element, str):
                wait = WebDriverWait(driver, 10)
                select_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, select_element)))
            elif not isinstance(select_element, WebElement):
                raise ValueError("select_element must be a string or WebElement")

            # Wait until the select element is interactable
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(select_element))

            # Ensure all options are loaded
            WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, "option")))

            select_obj = Select(select_element)

            try:
                current_selected = select_obj.first_selected_option.text
            except NoSuchElementException:
                current_selected = "None"

            self.logger.debug(f"{numero_contrat} * Selection actuelle dans le select: {current_selected}")

            select_obj.select_by_index(1)

            try:
                new_selected = select_obj.first_selected_option.text
            except NoSuchElementException:
                new_selected = "None"

            self.logger.info(f"{numero_contrat} * Nouveau Libellé d'établissement : {new_selected}")
        except TimeoutException:
            self.logger.error("Le sélecteur n'a pas pu être localisé ou n'était pas interactif dans le temps imparti.")
        except Exception as e:
            self.logger.error(f"{numero_contrat} * Erreur lors de la mise à jour de l'élément select: {e}")
    
    def select_role(self, select_element, role):
        try:
            select_obj = Select(select_element)
            for option in select_obj.options:
                if option.text.strip() == role:
                    select_obj.select_by_visible_text(role)
                    self.logger.debug(f"Rôle '{role}' sélectionné dans le sélecteur.")
                    return
            self.logger.error(f"Rôle '{role}' non trouvé dans le sélecteur.")
        except Exception as e:
            self.logger.error(f"Erreur lors de la sélection du rôle '{role}' : {e}")

    
    def extraire_valeurs_contrat(self, numero_contrat):
        """
        Charge le fichier Excel et extrait les valeurs de dépôt et de traitement pour un numéro de contrat donné.
        
        :param numero_contrat: Le numéro de contrat à rechercher.
        :param fichier_excel: Le chemin vers le fichier Excel.
        :return: Un tuple contenant les valeurs de dépôt et de traitement (ancien_depot_value, ancien_traitement_value).
        """
        try:
            # Charger le fichier Excel
            df = pd.read_excel('data/data_traitement/ROYE PIC - Transfert des contrats Affranchigo 070125 V2.xlsx')
            self.logger.debug("Fichier Excel chargé.")
            
            # Trouver la ligne correspondant au numéro de contrat
            row = df[df['Contrat Nb'] == numero_contrat]
            if row.empty:
                self.logger.error(f'Contrat {numero_contrat} introuvable dans le fichier Excel')
                return None, None

            # Extraire les valeurs de dépôt et de traitement
            ancien_depot_value = row['Code REGATE Dépôt actuel'].values[0] if pd.notna(row['Code REGATE Dépôt actuel'].values[0]) else None
            ancien_traitement_value = row['Code REGATE Traitement actuel'].values[0] if pd.notna(row['Code REGATE Traitement actuel'].values[0]) else None

            nouveau_depot_value = row['Nouveau code REGATE Dépôt'].values[0] if pd.notna(row['Nouveau code REGATE Dépôt'].values[0]) else None
            nouveau_traitement_value = row['Nouveau code REGATE Traitement'].values[0] if pd.notna(row['Nouveau code REGATE Traitement'].values[0]) else None

            ancien_depot_value = str(int(ancien_depot_value)) if isinstance(ancien_depot_value, float) else str(ancien_depot_value) if ancien_depot_value else None
            ancien_traitement_value = str(int(ancien_traitement_value)) if isinstance(ancien_traitement_value, float) else str(ancien_traitement_value) if ancien_traitement_value else None

            nouveau_depot_value = str(int(nouveau_depot_value)) if isinstance(nouveau_depot_value, float) else str(nouveau_depot_value) if nouveau_depot_value else None
            nouveau_traitement_value = str(int(nouveau_traitement_value)) if isinstance(nouveau_traitement_value, float) else str(nouveau_traitement_value) if nouveau_traitement_value else None

            return ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value

        except Exception as e:
            self.logger.error(f"Erreur lors de l'extraction des valeurs pour le contrat {numero_contrat} : {e}")
            return None, None
    
    def handle_case_premium(self, numero_contrat, dictionnaire):
        """Traitement principal pour chaque cas."""
        self.logger.info(f"{numero_contrat} * Traitement du contrat Affranchigo Premium")
        # Charger le fichier Excel
        df = pd.read_excel('data/data_traitement/ROYE PIC - Transfert des contrats Affranchigo 070125 V2.xlsx')
        self.logger.debug("Fichier Excel chargé.")

        # Trouver la ligne correspondant au numéro de contrat
        row = df[df['Contrat Nb'] == numero_contrat]
        if row.empty:
            self.logger.error(f'Contrat {numero_contrat} introuvable dans le fichier Excel')
            return

        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        time.sleep(2)
        elements = self.initialize_selectors(numero_contrat)
        time.sleep(5)
        if not elements:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
    
        current_value_first = None
        current_value_second = None

        # Récupérer l'élément et obtenir sa valeur si c'est un <input>
        current_value_first = elements['select_first_role'].get_attribute('value')
        self.logger.debug(f"Valeur actuelle Premier Rôle: {current_value_first}")

        current_value_second = elements['select_second_role'].get_attribute('value')
        self.logger.debug(f"Valeur actuelle Second Rôle: {current_value_second}")

        # Logique de traitement en fonction des valeurs de dépôt et de traitement
        if ancien_depot_value and ancien_traitement_value:
            if ancien_depot_value == ancien_traitement_value:
                if current_value_first == "Dépôt et Traitement":
                    self.dépôt_et_traitement_égaux(numero_contrat, dictionnaire)
                elif (current_value_first == "Dépôt" or current_value_first == "Traitement") and (current_value_second == "Traitement" or current_value_second == "Dépôt"):
                    self.dépôt_traitement_diff_egaux(numero_contrat, dictionnaire)
                elif (current_value_first == "Dépôt" or current_value_second == "Dépôt") and (current_value_second == "" or current_value_first == ""):
                    self.dépôt_only_value_egaux(numero_contrat, dictionnaire)
                elif (current_value_first == "Traitement" or current_value_second == "Traitement") and (current_value_second == "" or current_value_first == ""):
                    self.traitement_only_value_egaux(numero_contrat, dictionnaire)
            else:
                if current_value_first == "Dépôt et Traitement":
                    self.dépôt_et_traitement_non_égaux(numero_contrat, dictionnaire)
                elif (current_value_first == "Dépôt" or current_value_first == "Traitement") and (current_value_second == "Traitement" or current_value_second == "Dépôt"):
                    self.dépôt_traitement_diff_non_égaux(numero_contrat, dictionnaire)
                elif (current_value_first == "Dépôt" or current_value_second == "Dépôt") and (current_value_second == "" or current_value_first == ""):
                    self.dépôt_only_value_diff(numero_contrat, dictionnaire)
                elif (current_value_first == "Traitement" or current_value_second == "Traitement") and (current_value_second == "" or current_value_first == ""):
                    self.traitement_only_value_diff(numero_contrat, dictionnaire)
        elif ancien_depot_value:
            if current_value_first == "Dépôt et Traitement":
                self.dépôt_et_traitement_dépôt_modification(numero_contrat, dictionnaire)
            elif (current_value_first == "Dépôt" or current_value_first == "Traitement") and (current_value_second == "Traitement" or current_value_second == "Dépôt"):
                self.dépôt_diff_non_égaux(numero_contrat, dictionnaire)
            elif (current_value_first == "Dépôt" or current_value_second == "Dépôt") and (current_value_second == "" or current_value_first == ""):
                self.dépôt_only_depôt_only(numero_contrat, dictionnaire)
            elif (current_value_first == "Traitement" or current_value_second == "Traitement") and (current_value_second == "" or current_value_first == ""):
                self.traitement_only_depôt_only(numero_contrat, dictionnaire)
        elif ancien_traitement_value:
            if current_value_first == "Dépôt et Traitement":
                self.dépôt_et_traitement_traitement_modification(numero_contrat, dictionnaire)
            elif (current_value_first == "Dépôt" or current_value_first == "Traitement") and (current_value_second == "Traitement" or current_value_second == "Dépôt"):
                self.traitement_diff_non_égaux(numero_contrat, dictionnaire)
            elif (current_value_first == "Dépôt" or current_value_second == "Dépôt") and (current_value_second == "" or current_value_first == ""):
                self.dépôt_only_traitement_only(numero_contrat, dictionnaire)
            elif (current_value_first == "Traitement" or current_value_second == "Traitement") and (current_value_second == "" or current_value_first == ""):
                self.traitement_only_traitement_only(numero_contrat, dictionnaire)
        else:
            self.logger.error(f'Aucune valeur à traiter pour le contrat {numero_contrat}')

    
    # Ajout des méthodes spécifiques appelées dans handle_case_lib
    def dépôt_et_traitement_égaux(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors_radio_oui()
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_element = selectors.get("input_first_regate")
        select_element = selectors.get("select_first_etablissement")
        ancien_valeur_input_first_regate = self.get_selector_value(input_element)
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé 'Dépôt et Traitement'")
        self.logger.info(f"{numero_contrat} * Ancienne valeur Régate {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}")

        self.update_input(input_element, nouveau_traitement_value, numero_contrat)
        self.update_select_element(self.driver, select_element, numero_contrat)

        time.sleep(2)

        self.select_time_in_selectors(numero_contrat)
    
        self.logger.info(f"{numero_contrat} * Mise à jour du code régate. Enregistrement. Sortie.")
    
    def dépôt_traitement_diff_egaux(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors(numero_contrat)
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Traitement et un bloc Dépôt'")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}") 
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            try:
                radio_oui = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.ID, "g0_p265|0_c24954_v1"))
                )
            except TimeoutException:
                self.logger.error("Radio_oui non trouvé.")

            contrat_data = dictionnaire.get(numero_contrat, {})
            new_value_depot = contrat_data.get('Nouveau code REGATE Dépôt', '')

            if not new_value_depot:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            if not radio_oui.is_displayed() or not radio_oui.is_enabled():
                self.logger.debug(f"L'élément radio_non n'est pas visible ou activé pour le contrat {numero_contrat}")
                time.sleep(2)
                pass

            # Utilisation de JavaScript pour cliquer sur le bouton radio_non
            self.driver.execute_script("arguments[0].click();", radio_oui)
            self.logger.info(f"{numero_contrat} * Clic sur le bouton oui.")
            time.sleep(3)
            
            selectors = self.initialize_selectors_radio_oui()
            if not selectors:
                self.logger.error(f"Impossible d'initialiser les sélecteurs après le clic pour le contrat {numero_contrat}")
                return
            
            input_regate_first = selectors.get("input_first_regate")
            
            select_element_first = selectors.get("select_first_etablissement")
            
            select_role_first = selectors.get("select_first_role")
            
            if not all([input_regate_first, select_element_first, select_role_first]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            self.update_input(input_regate_first, new_value_depot, numero_contrat)
            
            self.update_select_element(self.driver, select_element_first, numero_contrat)

            time.sleep(2)

            self.select_time_in_selectors(numero_contrat)
            
            # Sélection des rôles
            """ Dans ce cas le rôle se met automatiquement en "Dépôt et Traitement" """
            
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")
    
    def dépôt_only_value_egaux(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Dépôt et un bloc san rôle ")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}") 
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            try:
                radio_oui = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.ID, "g0_p265|0_c24954_v1"))
                )
            except TimeoutException:
                self.logger.error("Radio_non no trouvé.")

            if not nouveau_depot_value:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            if not radio_oui.is_displayed() or not radio_oui.is_enabled():
                self.logger.debug(f"L'élément radio_non n'est pas visible ou activé pour le contrat {numero_contrat}")
                time.sleep(2)
                pass

            # Utilisation de JavaScript pour cliquer sur le bouton radio_non
            self.driver.execute_script("arguments[0].click();", radio_oui)
            self.logger.info(f"{numero_contrat} * Clic sur le bouton oui.")
            time.sleep(3)
            
            input_regate_first = selectors.get("input_first_regate")
            
            select_element_first = selectors.get("select_first_etablissement")

            select_role_first = selectors.get("select_first_role")

            if not all([input_regate_first, select_element_first, select_role_first]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            self.update_input(input_regate_first, nouveau_depot_value, numero_contrat)
            
            self.update_select_element(self.driver, select_element_first, numero_contrat)

            time.sleep(2)

            self.select_time_in_selectors(numero_contrat)
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")
    
    def traitement_only_value_egaux(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Traitement et un bloc sans rôle")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}") 
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            try:
                radio_oui = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.ID, "g0_p265|0_c24954_v1"))
                )
            except TimeoutException:
                self.logger.error("Radio_oui non trouvé.")

            ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)

            contrat_data = dictionnaire.get(numero_contrat, {})

            new_value_depot = contrat_data.get('Nouveau code REGATE Dépôt', '')

            if not new_value_depot:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            if not radio_oui.is_displayed() or not radio_oui.is_enabled():
                self.logger.debug(f"L'élément radio_non n'est pas visible ou activé pour le contrat {numero_contrat}")
                time.sleep(2)
                pass

            # Utilisation de JavaScript pour cliquer sur le bouton radio_non
            self.driver.execute_script("arguments[0].click();", radio_oui)
            self.logger.info(f"{numero_contrat} * Clic sur le bouton non.")
            time.sleep(3)
            
            input_regate_first = selectors.get("input_first_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_role_first = selectors.get("select_first_role")

            if not all([input_regate_first, select_element_first, select_role_first]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            self.update_input(input_regate_first, new_value_depot, numero_contrat)

            self.update_select_element(self.driver, select_element_first, numero_contrat)

            time.sleep(2)

            self.select_time_in_selectors(numero_contrat)

            # Sélection des rôles
            """Pas de choix de rôles pour ce traitement"""
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")

    def dépôt_et_traitement_non_égaux(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors_radio_oui()
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Traitement")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate None")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}") 
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            
            try:
                radio_non = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.ID, "g0_p265|0_c24954_v0"))
                )
            except TimeoutException:
                self.logger.error("Radio_non non trouvé.")
                return

            if not nouveau_traitement_value or not nouveau_depot_value:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            if not radio_non.is_displayed() or not radio_non.is_enabled():
                self.logger.debug(f"L'élément radio_non n'est pas visible ou activé pour le contrat {numero_contrat}")
                time.sleep(2)
                pass

            # Utilisation de JavaScript pour cliquer sur le bouton radio_non
            self.driver.execute_script("arguments[0].click();", radio_non)
            self.logger.info(f"{numero_contrat} * Clic sur le bouton non.")
            time.sleep(3)
            
            try:
                input_second_regate = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.ID, "g0_p265|0_r2055_c2056[0]"))
                )
            except TimeoutException:
                self.logger.error("Input second_regate non trouvé.")
                return

            input_second_regate.click()
            input_second_regate.send_keys(Keys.TAB)
            
            selectors = self.initialize_selectors(numero_contrat)
            if not selectors:
                self.logger.error(f"Impossible d'initialiser les sélecteurs après le clic pour le contrat {numero_contrat}")
                return
            
            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                return

            self.update_input(input_regate_first, nouveau_depot_value, numero_contrat)

            self.update_select_element(self.driver, select_element_first, numero_contrat)

            self.update_input(input_regate_second, nouveau_traitement_value, numero_contrat)

            self.update_select_element(self.driver, select_element_second, numero_contrat)

            time.sleep(2)
            
            # Sélection des rôles
            #self.select_role(select_role_first, "Dépôt")
            #self.select_role(select_role_second, "Traitement")
            self.select_time_in_selectors(numero_contrat)
            
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")
    
    def dépôt_traitement_diff_non_égaux(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Traitement et un bloc Dépôt")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}")

        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return

            if not nouveau_traitement_value or not nouveau_depot_value:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return
            
            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            self.update_input(input_regate_first, nouveau_depot_value, numero_contrat)

            self.update_select_element(self.driver, select_element_first, numero_contrat)

            self.update_input(input_regate_second, nouveau_traitement_value, numero_contrat)
            
            self.update_select_element(self.driver, select_element_second, numero_contrat)

            time.sleep(2)
            
            # Sélection des rôles
            #self.select_role(select_role_first, "Dépôt")
            #self.select_role(select_role_second, "Traitement")
            self.select_time_in_selectors(numero_contrat)
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")

    def dépôt_only_value_diff(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Dépôt et un bloc sans rôle")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}")
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            if not nouveau_traitement_value or not nouveau_depot_value:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            self.update_input(input_regate_first, nouveau_depot_value, numero_contrat)
            self.update_input(input_regate_second, nouveau_traitement_value, numero_contrat)
            
            self.update_select_element(self.driver, select_element_first, numero_contrat)
            self.update_select_element(self.driver, select_element_second, numero_contrat)

            time.sleep(2)
            

            # Sélection des rôles
            #self.select_role(select_role_first, "Dépôt")
            #self.select_role(select_role_second, "Traitement")
            self.select_time_in_selectors(numero_contrat)
            
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")

    def traitement_only_value_diff(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Traitement et un bloc sans rôle")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}")
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            
            contrat_data = dictionnaire.get(numero_contrat, {})
            new_value_traitement = contrat_data.get('Nouveau code REGATE Traitement', '')
            new_value_depot = contrat_data.get('Nouveau code REGATE Dépôt', '')

            if not new_value_traitement or not new_value_depot:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass
            # Vérification du titre de l'option sélectionnée pour select_role_first
            select_obj_first = Select(select_role_first)
            select_obj_second = Select(select_role_second)

            title_first_role = select_obj_first.first_selected_option.get_attribute('title')
            title_second_role = select_obj_second.first_selected_option.get_attribute('title')

            if title_first_role == "Traitement" :
                self.update_input(input_regate_first, new_value_traitement, numero_contrat)
                # MAJ de l'établissement
                self.update_select_element(self.driver, select_element_first, numero_contrat)
                # Vérification de la valeur de l'input_regate_first
                #self.select_role(select_role_first, "Traitement")

            elif title_first_role == "Dépôt" : 
                self.update_input(input_regate_first, new_value_depot, numero_contrat)
                # MAJ de l'établissement
                self.update_select_element(self.driver, select_element_first, numero_contrat)
                # Vérification de la valeur de l'input_regate_first
                #self.select_role(select_role_first, "Dépôt")

            elif title_second_role == "Traitement" :
                self.update_input(input_regate_second, new_value_traitement, numero_contrat)
                # MAJ de l'établissement
                self.update_select_element(self.driver, select_element_second, numero_contrat)
                # Vérification de la valeur de l'input_regate_first
                #self.select_role(select_role_second, "Traitement")
            
            elif title_second_role == "Dépôt" :
                self.update_input(input_regate_second, new_value_depot, numero_contrat)
                # MAJ de l'établissement
                self.update_select_element(self.driver, select_element_second,numero_contrat)
                # Vérification de la valeur de l'input_regate_first
                #self.select_role(select_role_second, "Dépôt")
   
            self.select_time_in_selectors()
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")

    def dépôt_et_traitement_dépôt_modification(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors_radio_oui()
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc 'Dépôt et Traitement' ")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate None")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            try:
                radio_non = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.ID, "g0_p265|0_c24954_v0"))
                )
            except TimeoutException:
                self.logger.error("Radio_non no trouvé.")

                # Extraire les valeurs de dépôt et de traitement depuis le fichier Excel
            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            if not all([input_regate_first, select_element_first, select_role_first]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass
            
            if not ancien_valeur_input_first_regate :
                self.logger.error(f'Valeurs du code premier code régate introuvable pour le numéro de contrat : {numero_contrat}')
                pass

            if not nouveau_depot_value:
                self.logger.error(f"Valeurs manquantes dans le fichier Excel pour le contrat {numero_contrat}")
                return

            if not radio_non.is_displayed() or not radio_non.is_enabled():
                self.logger.debug(f"L'élément radio_non n'est pas visible ou activé pour le contrat {numero_contrat}")
                time.sleep(2)
                pass

            # Utilisation de JavaScript pour cliquer sur le bouton radio_non
            self.driver.execute_script("arguments[0].click();", radio_non)
            self.logger.debug("Clic sur radio_non effectué avec JavaScript.")
            time.sleep(3)
            try:
                input_second_regate = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.ID, "g0_p265|0_r2055_c2056[0]"))
                )
            except TimeoutException:
                self.logger.error("Radio_non no trouvé.")
            input_second_regate.click()
            input_second_regate.send_keys(Keys.TAB)
            
            selectors = self.initialize_selectors(numero_contrat)
            if not selectors:
                self.logger.error(f"Impossible d'initialiser les sélecteurs après le clic pour le contrat {numero_contrat}")
                return
            
            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")
            
            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass
            # MAJ Nouvelle valeur Dépôt
            self.update_input(input_regate_first, ancien_valeur_input_first_regate,numero_contrat)
            
            self.update_select_element(self.driver, select_element_first, numero_contrat)
            # MAJ nouvelle valeur Dépôt
            self.update_input(input_regate_second, nouveau_depot_value, numero_contrat)

            self.update_select_element(self.driver, select_element_second, numero_contrat)

            time.sleep(2)


            # Sélection des rôles
            #self.select_role(select_role_first, "Dépôt")
            #self.select_role(select_role_second, "Traitement")
            self.select_time_in_selectors(numero_contrat)
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")
    
    def dépôt_diff_non_égaux(self, numero_contrat, dictionnaire): 
        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Traitement et un bloc Dépôt")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            
            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            if not nouveau_depot_value:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")
    
            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            # Vérification du titre de l'option sélectionnée pour select_role_first
            select_obj_first = Select(select_role_first)
            select_obj_second = Select(select_role_second)

            title_first_role = select_obj_first.first_selected_option.get_attribute('title')
            title_second_role = select_obj_second.first_selected_option.get_attribute('title')
            
            if title_first_role == "Dépôt":

                self.logger.debug(f"Modification de dépôt: de {ancien_valeur_input_first_regate} à {nouveau_depot_value}")

                self.update_input(input_regate_first, nouveau_depot_value, numero_contrat)

                self.update_select_element(self.driver, select_element_first, numero_contrat)

            elif title_second_role == "Dépôt":

                self.logger.debug(f"Modification de dépôt: de {ancien_valeur_input_first_regate} à {nouveau_depot_value}")

                self.update_input(input_regate_second, nouveau_depot_value ,numero_contrat)

                self.update_select_element(self.driver, select_element_second, numero_contrat)
            else:
                self.logger.error(f"Valeur actuelle de dépôt {input_regate_first.get_attribute('value')} ne correspond pas à la valeur cible {ancien_valeur_input_first_regate}")

            time.sleep(2)

            self.select_time_in_selectors(numero_contrat)
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")
    
    def dépôt_only_depôt_only(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Dépôt et un bloc sans rôle")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return

            
            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            if not nouveau_depot_value:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return
            # Vérification du titre de l'option sélectionnée pour select_role_first
            select_obj_first = Select(select_role_first)
            select_obj_second = Select(select_role_second)

            title_first_role = select_obj_first.first_selected_option.get_attribute('title')
            title_second_role = select_obj_second.first_selected_option.get_attribute('title')

            if title_first_role == "Dépôt":

                self.logger.debug(f"Modification de dépôt: de {ancien_valeur_input_first_regate} à {nouveau_depot_value}")

                self.update_input(input_regate_first, nouveau_depot_value, numero_contrat)

                self.update_select_element(self.driver, select_element_first, numero_contrat)

            elif title_second_role == "Dépôt":

                self.logger.debug(f"Modification de dépôt: de {ancien_valeur_input_first_regate} à {nouveau_depot_value}")

                self.update_input(input_regate_second, nouveau_depot_value, numero_contrat)

                self.update_select_element(self.driver, select_element_second, numero_contrat)
            else:
                self.logger.error(f"Valeur actuelle de dépôt {nouveau_depot_value} ne correspond pas à la valeur cible {ancien_valeur_input_first_regate}")

            time.sleep(2)

            
            if not select_role_second.get_attribute("value"): 
                # Sélection des rôles
                #self.select_role(select_role_second, "Traitement")
                self.driver.save_screenshot(f"{numero_contrat} * absence_traitement.png")

            #self.select_role(select_role_first, "Dépôt")
            self.select_time_in_selectors(numero_contrat)
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")
    
    def traitement_only_depôt_only(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Traitement et un bloc sans rôle")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel Dépôt {nouveau_depot_value}")
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            
            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            if not nouveau_depot_value:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            # Vérification du titre de l'option sélectionnée pour select_role_first
            select_obj_first = Select(select_role_first)
            select_obj_second = Select(select_role_second)

            title_first_role = select_obj_first.first_selected_option.get_attribute('title')
            title_second_role = select_obj_second.first_selected_option.get_attribute('title')

            if title_first_role == "Dépôt":

                self.logger.debug(f"Modification de dépôt: de {ancien_valeur_input_first_regate} à {nouveau_depot_value}")

                self.update_input(input_regate_first, nouveau_depot_value, numero_contrat)

                self.update_select_element(self.driver, select_element_first, numero_contrat)

            elif title_second_role == "Dépôt":

                self.logger.debug(f"Modification de dépôt: de {ancien_valeur_input_first_regate} à {nouveau_depot_value}")

                self.update_input(input_regate_second, nouveau_depot_value, numero_contrat)

                self.update_select_element(self.driver, select_element_second, numero_contrat)
            else:
                self.logger.error(f"{numero_contrat} * Valeur actuelle de dépôt {nouveau_depot_value} ne correspond pas à la valeur cible {ancien_valeur_input_first_regate}")

            time.sleep(2)

            
            if not select_role_first.get_attribute("value"): 
                # Sélection des rôles
                #self.select_role(select_role_second, "Dépôt")
                self.driver.save_screenshot(f"{numero_contrat} * absence_depôt.png")
            # Sélection des rôles
            #self.select_role(select_role_second, "Traitement")
            self.select_time_in_selectors(numero_contrat)
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")

    def dépôt_et_traitement_traitement_modification(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors_radio_oui()
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_element = selectors.get("input_first_regate")
        select_element = selectors.get("select_first_etablissement")
        ancien_valeur_input_first_regate = self.get_selector_value(input_element)
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé 'Dépôt et Traitement'")
        self.logger.info(f"{numero_contrat} * Ancienne valeur Régate {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}")
        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            
            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")


            if not all([input_regate_first, select_element_first, select_role_first]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass
    
            if not ancien_valeur_input_first_regate :
                self.logger.error(f'Valeurs du code premier code régate introuvable pour le numéro de contrat : {numero_contrat}')
                pass

            if not nouveau_traitement_value:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            try:
                input_second_regate = WebDriverWait(self.driver, 30).until(
                    EC.presence_of_element_located((By.ID, "#g0_p10858\\|0_r72238_c72239\\[1\\]"))
                )
            except TimeoutException:
                self.logger.error("Radio_non no trouvé.")
            input_second_regate.click()
            input_second_regate.send_keys(Keys.TAB)
            
            
            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")
            
            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(5)
                pass
            
            self.update_input(input_regate_first, ancien_valeur_input_first_regate,numero_contrat)
            
            self.update_select_element(self.driver, select_element_first, numero_contrat)

            self.update_input(input_second_regate, nouveau_traitement_value, numero_contrat)

            self.update_select_element(self.driver, select_element_second, numero_contrat)

            time.sleep(2)


            # Sélection des rôles
            #self.select_role(select_role_first, "Dépôt")
            #self.select_role(select_role_second, "Traitement")
            self.select_time_in_selectors(numero_contrat)
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")
    

    def traitement_diff_non_égaux(self, numero_contrat, dictionnaire):
        self.logger.info(f"{numero_contrat} * 'traitement_diff_non_égaux' Rôle à l'arrivée un bloc Traitement et un bloc Dépôt")

        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return

        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)

        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)

        self.logger.info(f"{numero_contrat} * Ancienne première valeur Régate {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne seconde valeur Régate {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}")

        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return

            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            # Vérification de la présence des sélecteurs
            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                return

            if not nouveau_traitement_value:
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            # Attendre pour s'assurer que les éléments sont chargés
            time.sleep(2)

            # Récupérer les titres des sélecteurs
            title_first_role = select_role_first.get_attribute('title')
            title_second_role = select_role_second.get_attribute('title')

            # Mettre à jour en fonction des rôles
            if title_first_role == "Traitement":
                self.logger.debug(f"Modification de traitement: de {ancien_valeur_input_first_regate} à {nouveau_traitement_value}")
                self.update_input(input_regate_first, nouveau_traitement_value, numero_contrat)
                self.update_select_element(self.driver, select_element_first, numero_contrat)

            elif title_second_role == "Traitement":
                self.logger.debug(f"Modification de traitement: de {ancien_valeur_input_first_regate} à {nouveau_traitement_value}")
                self.update_input(input_regate_second, nouveau_traitement_value, numero_contrat)
                self.update_select_element(self.driver, select_element_second, numero_contrat)
            else:
                self.logger.debug(f"Le titre du sélecteur n'est pas 'Traitement', il est '{title_first_role}'")
                self.logger.error(f"Valeur actuelle de dépôt {title_first_role} ne correspond pas à la valeur cible {ancien_valeur_input_first_regate}")

            # Attente pour appliquer les modifications
            time.sleep(2)

            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.select_time_in_selectors(numero_contrat)
            self.submit_liberte(numero_contrat)

        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_oui.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")


    def dépôt_only_traitement_only(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Dépôt et un bloc sans rôle")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}")

        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return
            
            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            if not nouveau_traitement_value :
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return
            
            # Vérification du titre de l'option sélectionnée pour select_role_first
            select_obj_first = Select(select_role_first)
            select_obj_second = Select(select_role_second)

            title_first_role = select_obj_first.first_selected_option.get_attribute('title')
            title_second_role = select_obj_second.first_selected_option.get_attribute('title')

            if title_first_role == "Traitement":

                self.logger.debug(f"Modification de dépôt: de {ancien_valeur_input_first_regate} à {nouveau_traitement_value}")

                self.update_input(input_regate_first, nouveau_traitement_value, numero_contrat)

                self.update_select_element(self.driver, select_element_first, numero_contrat)

            elif title_second_role == "Traitement":

                self.logger.debug(f"Modification de dépôt: de {ancien_valeur_input_first_regate} à {nouveau_traitement_value}")

                self.update_input(input_regate_second, nouveau_traitement_value, numero_contrat)

                self.update_select_element(self.driver, select_element_second, numero_contrat)
            else:
                self.logger.error(f"Valeur actuelle de dépôt {input_regate_first.get_attribute('value')} ne correspond pas à la valeur cible {ancien_valeur_input_first_regate}")
            
            time.sleep(2)
            

            # Sélection des rôles
            #self.select_role(select_role_first, "Dépôt")
            #self.select_role(select_role_second, "Traitement")
            
            self.select_time_in_selectors(numero_contrat)
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")
    
    def traitement_only_traitement_only(self, numero_contrat, dictionnaire):
        selectors = self.initialize_selectors(numero_contrat)
        if not selectors:
            self.logger.error(f"Impossible d'initialiser les sélecteurs pour le contrat {numero_contrat}")
            return
        input_regate_first = selectors.get("input_first_regate")
        input_regate_second = selectors.get("input_second_regate")
        ancien_depot_value, ancien_traitement_value, nouveau_depot_value, nouveau_traitement_value = self.extraire_valeurs_contrat(numero_contrat)
        ancien_valeur_input_first_regate = self.get_selector_value(input_regate_first)
        ancien_valeur_input_second_regate = self.get_selector_value(input_regate_second)
        self.logger.info(f"{numero_contrat} * Rôle à l'arrivé un bloc Dépôt et un bloc sans rôle")
        self.logger.info(f"{numero_contrat} * Ancienne  première valeur Régate  {ancien_valeur_input_first_regate}")
        self.logger.info(f"{numero_contrat} * Ancienne  seconde valeur Régate  {ancien_valeur_input_second_regate}")
        self.logger.info(f"{numero_contrat} * Futur valeur du fichier Excel traitement {nouveau_traitement_value}")

        try:
            if not isinstance(numero_contrat, str):
                self.logger.error(f"Le numero_contrat doit être une chaîne de caractères, reçu: {type(numero_contrat)}")
                return

            input_regate_first = selectors.get("input_first_regate")
            input_regate_second = selectors.get("input_second_regate")
            select_element_first = selectors.get("select_first_etablissement")
            select_element_second = selectors.get("select_second_etablissement")
            select_role_first = selectors.get("select_first_role")
            select_role_second = selectors.get("select_second_role")

            if not all([input_regate_first, input_regate_second, select_element_first, select_element_second, select_role_first, select_role_second]):
                self.logger.error(f"Un ou plusieurs sélecteurs sont manquants après le clic sur radio_non pour le contrat {numero_contrat}")
                time.sleep(3)
                pass

            if not nouveau_traitement_value :
                self.logger.error(f"Valeurs manquantes pour le contrat {numero_contrat}")
                return

            # Vérification du titre de l'option sélectionnée pour select_role_first
            select_obj_first = Select(select_role_first)
            select_obj_second = Select(select_role_second)

            title_first_role = select_obj_first.first_selected_option.get_attribute('title')
            title_second_role = select_obj_second.first_selected_option.get_attribute('title')

            if title_first_role == "Traitement":

                self.logger.debug(f"Modification de dépôt: de {ancien_valeur_input_first_regate} à {nouveau_traitement_value}")

                self.update_input(input_regate_first, nouveau_traitement_value, numero_contrat)

                self.update_select_element(self.driver, select_element_first, numero_contrat)

            elif title_second_role == "Traitement":

                self.logger.debug(f"Modification de dépôt: de {ancien_valeur_input_first_regate} à {nouveau_traitement_value}")

                self.update_input(input_regate_second, nouveau_traitement_value, numero_contrat)

                self.update_select_element(self.driver, select_element_second, numero_contrat)
            else:
                self.logger.error(f"Valeur actuelle de dépôt {nouveau_traitement_value} ne correspond pas à la valeur cible {ancien_valeur_input_first_regate}")
            
            time.sleep(2)
            

            # Sélection des rôles
            #self.select_role(select_role_first, "Dépôt")
            #self.select_role(select_role_second, "Traitement")
            
            self.select_time_in_selectors(numero_contrat)
            self.logger.debug(f"{numero_contrat} * Mise à jour des inputs et sélecteurs effectuée.")
            self.submit_liberte(numero_contrat)
        except TimeoutException:
            self.logger.error("Timeout lors de la tentative de clic sur radio_non.")
        except Exception as e:
            self.logger.error(f"Erreur lors du traitement du contrat {numero_contrat} : {e}")

