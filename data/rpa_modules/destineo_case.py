from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.remote.webelement import WebElement

from rpa_modules.debug import setup_logger


class DestineoCase:
    def __init__(self, driver, pool, logger=None):
        """
        Initialise la classe avec un WebDriver, un pool, et un logger.
        :param driver: L'instance de WebDriver à utiliser.
        :param pool: Le pool de WebDrivers partagé.
        :param logger: Logger pour les logs. Un logger par défaut est créé si non fourni.
        """
        self.driver = driver  # Initialisation du WebDriver
        self.pool = pool
        self.logger = logger or setup_logger("destineo_case.log")
        
    def submit_destineo(self, driver):
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
            wait = WebDriverWait(driver, 20)
            
            selectors["adresse_facturation_EL"] = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#cptLeft > div:nth-child(13) select"))
            )
            selectors["input_first_regate_EL_kdo"] = wait.until(
                EC.presence_of_element_located((By.ID, "g0_p143|0_r442_c443"))
            )
    
            selectors["select_first_etablissement_EL"] = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#g0_p143\\|0_r442\\[0\\] > div > critere-form:nth-child(3) > div.form-group.critere_psc > input-etb-prest > div > select"))
            )
            
            self.logger.debug("Tous les sélecteurs ont été initialisés avec succès.")
        except TimeoutException as e:
            self.logger.error(f"Timeout lors de l'initialisation des sélecteurs: {e}")
        except NoSuchElementException as e:
            self.logger.error(f"Élément non trouvé lors de l'initialisation des sélecteurs: {e}")
        except Exception as e:
            self.logger.error(f"Erreur lors de l'initialisation des sélecteurs: {e}")
        return selectors
        
    def update_input(self, element, new_value):
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
        self.logger.info("Nouveau code régate attribué")

    def update_select_element(self, driver, select_element):
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

            self.logger.debug(f"Selection actuelle dans le select: {current_selected}")

            select_obj.select_by_index(1)

            try:
                new_selected = select_obj.first_selected_option.text
            except NoSuchElementException:
                new_selected = "None"

            self.logger.debug(f"Index {1} sélectionné dans le select, nouvelle selection: {new_selected}")
            self.logger.info("Nouvel établissement attribué.")
        except TimeoutException:
            self.logger.error("Le sélecteur n'a pas pu être localisé ou n'était pas interactif dans le temps imparti.")
        except Exception as e:
            self.logger.error(f"Erreur lors de la mise à jour de l'élément select: {e}")
    
    def traitement_destineo_el_kdo(self, driver, wait, numero_contrat, dictionnaire):
        self.logger.debug('Debut du traitement pour le contrat numero: {}'.format(numero_contrat))
        selectors = self.initialize_selectors(driver)

        if not selectors:
            self.logger.debug('Aucun selecteur initialise pour le contrat numero: {}'.format(numero_contrat))
            return

        input_element = selectors.get("input_first_regate_EL_kdo")
        if not input_element:
            self.logger.debug('element input non trouve, contrat numero: {}'.format(numero_contrat))
            return
        select_element = selectors.get("select_first_etablissement_EL")

        adresse_facturation_el = selectors.get("adresse_facturation_EL")
        if adresse_facturation_el:
            adresse_facturation = Select(adresse_facturation_el)
            selected_option = adresse_facturation.first_selected_option
            if selected_option and selected_option.get_attribute('value'):
                self.logger.debug(f"Option selectionnee pour l'adresse de facturation: {selected_option.text}")
            else:
                # Aucune adresse de facturation selectionnee, on selectionne la première option disponible
                self.logger.debug("Aucune adresse de facturation selectionnee, selection de la premiere option disponible.")
                if adresse_facturation.options: 
                    adresse_facturation.select_by_index(0)
                    self.logger.debug(f"Premiere option selectionnee : {adresse_facturation.first_selected_option.text}")
                else:
                    self.logger.debug("Aucune option disponible dans le selecteur d'adresse de facturation.")
        else:
            self.logger.debug("Element select pour l adresse de facturation non trouve.")

        
        new_value = dictionnaire.get(numero_contrat, {}).get('Nouveau Régate Dépôt', '')
        # Mise à jour des entrees et selection
        self.update_input(input_element, new_value)
        # Si erreur initier le Select dans la fonction
        self.update_select_element(self.driver, select_element)
        self.logger.info("Traitement complet pour le contrat numero: {}".format(numero_contrat))
        # Envois du formulaire
        #self.submit_destineo(driver)
        self.logger.debug("Contrat traité")
    
    
    def handle_case_destineo(self, driver, wait, numero_contrat, dictionnaire):
        self.logger.debug("Debut du cas Destineo")
        try:
            h1_text = driver.find_element(By.TAG_NAME, "h1").text
            self.logger.debug(f"Texte de l en tete H1 : {h1_text}")
            if "Destineo esprit libre" in h1_text:
                self.logger.debug("Cas Destineo EL")
                self.traitement_destineo_el_kdo(driver, wait, numero_contrat, dictionnaire)

            elif "Destineo kdo" in h1_text:
                self.logger.debug("Cas Destineo KDO")
                self.traitement_destineo_el_kdo(driver, wait, numero_contrat, dictionnaire)
            
            elif "Destineo monde primo" in h1_text:
                self.logger.debug("Cas Destineo monde primo")
            
            elif "Destineo monde volume" in h1_text:
                self.logger.debug("Cas Destineo monde volume")
        except Exception as e:
            self.logger.exception(f"Service non reconnu : {e}")
            self.self.logger.error("Erreur_service_non_reconnu")
