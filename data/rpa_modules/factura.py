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
from queue import Queue, Empty
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from threading import Thread
# Importation du setup_logger uniquement
from concurrent.futures import ThreadPoolExecutor, as_completed
from rpa_modules.debug import setup_logger
from rpa_modules.data_processing import extract_contrat_numbers_to_json

load_dotenv()

STOP_FLAG = False

class FacturaRPA:
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
        self.url = "https://opencell-blg-sacc-auth.paas-04.build.net.intra.laposte.fr/auth/realms/opencell/protocol/openid-connect/auth?client_id=opencell-portal&redirect_uri=https%3A%2F%2Fopencell-blg-sacc.paas-04.build.net.intra.laposte.fr%2Fopencell%2Ffrontend%2Fdefault%2Fportal%2F&state=2cc467e8-591d-4627-8398-eca6c3827aa6&response_mode=fragment&response_type=code&scope=openid&nonce=f9923c51-6623-48e4-bdd4-fc9c9aced731"

        def wait_for_complete_redirection_and_log(self, driver, wait, numero_contrat, timeout=20):
            self.logger.debug("Attente de la redirection...")

            try:
                input_mot_de_passe = wait.until(EC.element_to_be_clickable((By.ID, "AUTHENTICATION.PASSWORD")))
                input_mot_de_passe.clear()
                wait.until(EC.element_to_be_clickable((By.ID, "AUTHENTICATION.PASSWORD")))
                input_mot_de_passe.send_keys(mot_de_passe)
                input_mot_de_passe.send_keys(Keys.RETURN)
                button_validation = wait.until(EC.element_to_be_clickable((By.ID, "validateButton")))
                button_validation.click()
                self.logger.debug("L'identification à été compléter.")
            except TimeoutException as e:
                self.logger.exception("La redirection ou le chargement de la page n'a pas été complet dans le temps imparti, ou l'élément cible n'a pas été trouvé.")
                self.save_non_modifiable(numero_contrat)
            except NoSuchElementException as e:
                self.logger.exception("L'Input mot de passe n'as pas été trouvé")
                self.save_non_modifiable(numero_contrat)
            except Exception as e:
                self.logger.debug("Erreur inattendue lors de l'attente de la redirection ou du chargement de la page.")
                self.save_non_modifiable(numero_contrat)

        def login(self, driver, wait, numero_contrat):
            self.logger.info("Tentative de connexion...")
            try:
                button_GAIA = wait.until(EC.element_to_be_clickable((By.ID, "social-oidc")))
                button_GAIA.click()
                self.logger.info("Click réussi sur GAIA")
                self.wait_for_complete_redirection_and_log(driver, wait, numero_contrat)
            except Exception as e:
                self.logger.error("Error lors de la connexion à GAIA")
        
        def selection_extracted(driver, wait, numero_contrat):
            try:
                icon_operation =  wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#main-content > div > div > div > div > div:nth-child(4) > a")))
                icon_operation.click()
                self.logger.debug("Clique sur Opération abouti.")
                span_request =  wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#root > div:nth-child(1) > div > div > main > div.MuiDrawer-root.MuiDrawer-docked.jss136 > div > div > div.jss203 > div.MenuItemsContainer.jss204 > div:nth-child(6)")))
                span_request.click()
                generateur_requete = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#root > div:nth-child(1) > div > div > main > div.MuiDrawer-root.MuiDrawer-docked.jss136 > div > div > div.jss203 > div.MenuItemsContainer.jss204.jss205 > a:nth-child(1)")))
                generateur_requete.click()
            except Exception as e:
                self.logger.error("Erreur Clique sur Opération.")





