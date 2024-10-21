from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webelement import WebElement

from rpa_modules import setup_logger

class ProxicompteCase:
    def __init__(self, driver, pool, logger=None):
        """
        Initialise la classe avec un WebDriver, un pool, et un logger.
        :param driver: L'instance de WebDriver à utiliser.
        :param pool: Le pool de WebDrivers partagé.
        :param logger: Logger pour les logs. Un logger par défaut est créé si non fourni.
        """
        self.driver = driver  # Initialisation du WebDriver
        self.pool = pool
        self.logger = logger or setup_logger("proxicompte_case.log")

    def is_select_value_present(self, select_selector):
        """Vérifie si un élément select a une valeur sélectionnée."""
        try:
            select_element = Select(
                self.driver.find_element(By.CSS_SELECTOR, select_selector)
            )
            return select_element.first_selected_option.text.strip() != ""
        except NoSuchElementException:
            return False

    def is_specific_option_selected(self, select_selector, option_text):
        """Vérifie si une option spécifique est sélectionnée dans un élément select."""
        try:
            select_element = Select(
                self.driver.find_element(By.CSS_SELECTOR, select_selector)
            )
            selected_option_text = select_element.first_selected_option.text.strip()
            return selected_option_text == option_text
        except NoSuchElementException:
            self.logger.debug(f"Selecteur '{select_selector}' non trouvé.")
            return False
    
    def submit_proxicompte(self, driver):
        # Trouver et cliquer sur le bouton de soumission
            try:
                submit_button = WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, "#odcFormCPV > button")
                    )
                )
                submit_button.click()

                WebDriverWait(driver, 10).until(
                    EC.url_changes(
                        "https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/configurateur/put-contract"
                    )
                )
                self.logger.info("Formulaire soumis avec succes.")
            
                try:
                    # Retour à l'URL de depart
                    url_de_depart = "https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/index/acces-dc"
                    driver.get(url_de_depart)
                    self.logger.info("Retour à l'URL de depart reussi.")
                except Exception as e:
                    self.logger.critical(f"Erreur lors de la navigation vers l'URL de depart : {e}")
                    driver.save_screenshot("debug_screenshot_erreur_retour.png")
            except TimeoutException:
                self.logger.error("Le bouton de soumission n'as pas ete trouve dans les temps")
                driver.save_screenshot("debug_screenshot_erreur_clic.png")
            except Exception as e:
                self.logger.critical(f"Erreur lors de la soumission formulaire : {e}")
                driver.save_screenshot("debug_screenshot_erreur_soumission_proxicompte.png")
    
    def initialize_selectors(self, driver):
        selectors = {}
        try:
            wait = WebDriverWait(driver, 10)
            selectors["adresse_facturation"] = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#cptLeft > div:nth-child(2) select")))
            selectors["input_first_regate"] = wait.until(EC.presence_of_element_located((By.ID, "g0_p10003|0_r6204_c6205")))
            selectors["select_first_etablissement"] = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#g0_p10003\\|0_r6204 > div > critere-form:nth-child(3) select")))
        except Exception as e:
            self.logger.error(f"Erreur lors de l'initialisation des selecteurs: {e}")
        return selectors
    
    def compare_input_values(self, driver, selectors, current_value, expected_value, numero_contrat):
        if str(current_value) != str(expected_value):
            # Creation locale des objets Select
            select_first_etablissement = Select(selectors["select_first_etablissement"])
            select_adresse_facturation = Select(selectors["adresse_facturation"])

            # Verifie que les objets crees sont des instances de Select
            if not isinstance(select_first_etablissement, Select) or not isinstance(select_adresse_facturation, Select) :
                self.logger.error("Les elements fournis ne sont pas des elements de selection valides.")
                return False

            # Recuperation des valeurs à partir des elements Select
            etablissement_value_first = select_first_etablissement.first_selected_option.get_attribute('title') if select_first_etablissement else "Non trouve"
            select_adresse_facturation = select_adresse_facturation.first_selected_option.get_attribute('title') if select_adresse_facturation else "Non trouve"

            # Creation du message d'erreur sous forme de dictionnaire
            error_info = {
                "message": "Erreur de validation pour le contrat",
                "numero_contrat": numero_contrat,
                "adresse_facturation": select_adresse_facturation,
                "valeur_attendue": expected_value,
                "valeur_trouvee": current_value,
                "premier_etablissement": etablissement_value_first
            }
            self.logger.error(error_info)

            return False

        return True
    
    def update_input(self, driver, element, new_value):
        self.logger.debug(f"Tentative de mise à jour de l'input avec la nouvelle valeur: {new_value}")
        element.clear()
        element.send_keys(new_value)
        element.send_keys(Keys.TAB)
        self.logger.info("Nouveau code regate attribue")

    def update_select_element(self, driver,  select_element, index):
        if isinstance(select_element, str):
            wait = WebDriverWait(driver, 10)
            
            # Attendre que le select_element soit présent
            select_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, select_element)))
        elif not isinstance(select_element, WebElement):
            raise ValueError("select_element must be a string or WebElement")

        # Attendre que les options soient chargées (2 secondes)
        WebDriverWait(driver, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "option")))
        
        select_obj = Select(select_element)
        
        try:
            current_selected = select_obj.first_selected_option.text
        except NoSuchElementException:
            current_selected = "None"
        
        self.logger.debug(f"Selection actuelle dans le select: {current_selected}")
        
        select_obj.select_by_index(index)
        
        try:
            new_selected = select_obj.first_selected_option.text
        except NoSuchElementException:
            new_selected = "None"
        
        self.logger.debug(f"Index {index} selectionne dans le select, nouvelle selection: {new_selected}")
        self.logger.info("Nouvel etablissement attribue.")
    
    def process_traitement(self, driver, numero_contrat, dictionnaire):
        self.logger.debug('Debut du traitement pour le contrat numero: {}'.format(numero_contrat))
        
        # Initialisation des selecteurs
        elements = self.initialize_selectors(driver)
        selectors = self.initialize_selectors(driver)
        
        if not selectors:
            self.logger.debug('Aucun selecteur initialise pour le contrat numero: {}'.format(numero_contrat))
            return

        input_element = elements.get("input_first_regate")
        if not input_element:
            self.logger.debug('element input non trouve, contrat numero: {}'.format(numero_contrat))
            return

        # Recuperation des valeurs courantes et attendues
        current_value = input_element.get_attribute('value')
        expected_value = dictionnaire.get(numero_contrat, {}).get('Ancien Régate Traitement', '')
        self.logger.debug('Valeur courante recuperee: "{}", Valeur attendue: "{}" pour le contrat numero: {}'.format(current_value, expected_value, numero_contrat))

        # Comparaison des valeurs
        if not self.compare_input_values(driver, selectors, current_value, expected_value, numero_contrat):
            self.logger.debug('echec de la comparaison des valeurs pour le contrat numero: {}'.format(numero_contrat))
            return
        # Mise à jour des entrees et selection
        self.update_input(driver, input_element, dictionnaire[numero_contrat].get('Nouveau Régate Traitement', ''))
        # Si erreur initier le Select dans la fonction
        self.update_select_element(selectors["select_first_etablissement"], 1)
        self.logger.info("Traitement complet pour le contrat numero: {}".format(numero_contrat))
    
    def handle_case_proxicompte(self, driver, numero_contrat, dictionnaire):
        self.logger.debug("Debut de Cas Proxicompte")
        elements = self.initialize_selectors(driver)
        select_first_element_role = Select(elements['adresse_facturation'])
        # Verifie les selecteurs
        if select_first_element_role is None :
            self.logger.error("Aucune adresse de facturation n est selectionner")
        elif elements["input_first_regate"] is None :
            self.logger.error("Code lieu de depot sans valeur")
        else :
            self.process_traitement(driver, numero_contrat, dictionnaire)
            self.logger.debug("Fin du traitement proxicompte")

            return 