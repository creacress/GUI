import os
import time
import threading
from queue import Queue
import psutil
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from rpa_modules.debug import setup_logger
from concurrent.futures import ThreadPoolExecutor, as_completed

class WebDriverPool:
    def __init__(self, initial_size=1, max_size=20, idle_timeout=300, logger=None):
        """
        Pool de WebDrivers avec auto-ajustement dynamique de la taille du pool.
        :param initial_size: Taille initiale du pool (pré-chargement).
        :param max_size: Taille maximale du pool.
        :param idle_timeout: Temps d'inactivité maximum d'un WebDriver avant d'être recyclé.
        """
        self.pool = Queue(maxsize=max_size)
        self.max_size = max_size
        self.current_size = 0
        self.idle_timeout = idle_timeout
        self.lock = threading.Lock()
        self.logger = logger or setup_logger('WebdriverPool.log')
        self.logger.debug("WebDriverPool initialized with max_size=%d", max_size)

        # Pré-charger les instances initiales de WebDriver
        for _ in range(initial_size):
            driver = self.create_driver()
            self.pool.put(driver)
            self.current_size += 1

    def create_driver(self):
        """
        Crée une nouvelle instance de WebDriver configurée avec les options requises
        et navigue directement vers l'URL fournie.
        """
        try:
            self.logger.debug("Creating a new WebDriver instance")

            driver_path = r'data/driver/msedgedriver.exe'

            service = Service(driver_path)
            options = Options()
            options.add_argument("--headless")  # Si tu veux rester en mode headless
            options.add_argument("--disable-gpu")
            options.add_argument("--disable-software-rasterizer")
            options.add_argument('--ignore-certificate-errors')
            options.add_argument('--ignore-ssl-errors')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument("--log-level=3")
            options.add_argument("--disable-extensions")
            options.add_argument("--disable-popup-blocking")
            options.add_argument("--disable-software-rasterizer")

            driver = webdriver.Edge(service=service, options=options)

            # Naviguer vers l'URL directement après la création
            driver.get("https://portail.e-facture.net/saml/saml-login.php?nomSP=ARTIMON_PROD")
            driver.last_used_time = time.time()

            self.logger.debug("WebDriver instance created and navigated to URL successfully")
            return driver
        except Exception as e:
            self.logger.error(f"Erreur lors de la création du WebDriver: {e}")
            raise


    def get_driver(self, url=None):
        with self.lock:
            if not self.pool.empty():
                driver = self.pool.get()
                self.logger.debug("Réutilisation d'une instance WebDriver existante.")
                try:
                    driver.execute_script("return document.readyState")
                except Exception as e:
                    self.logger.warning(f"WebDriver inactif ou corrompu, recréation d'un WebDriver : {e}")
                    driver = self.create_driver()
            else:
                if self.current_size < self.max_size:
                    driver = self.create_driver()
                    self.current_size += 1
                    self.logger.debug(f"Nouvelle instance WebDriver créée (taille actuelle : {self.current_size})")
                else:
                    self.logger.warning("Taille maximale du pool WebDriver atteinte, attente d'un WebDriver disponible.")

            if url:
                driver.get(url)
            driver.last_used_time = time.time()
            return driver


    def return_driver(self, driver):
        with self.lock:
            if driver:
                try:
                    # Naviguer vers l'URL de départ avant de remettre le driver dans le pool
                    start_url = "https://portail.e-facture.net/saml/saml-login.php?nomSP=ARTIMON_PROD"
                    self.logger.debug(f"Retour à l'URL de départ {start_url} avant de retourner le WebDriver au pool.")
                    
                    driver.get(start_url)

                    # Vérifie si le WebDriver est encore actif en utilisant un script basique
                    status = driver.execute_script("return document.readyState")
                    if status != "complete":
                        raise Exception(f"WebDriver status non 'complete', status actuel : {status}")
                    
                    # Remettre le WebDriver dans le pool
                    self.pool.put(driver)
                    self.logger.debug(f"WebDriver remis dans le pool. Taille actuelle du pool : {self.pool.qsize()}")

                except Exception as e:
                    # Si le WebDriver est inactif ou s'il y a une erreur, ferme-le et décrémente le pool
                    self.logger.warning(f"WebDriver inactif ou erreur détectée, suppression du driver: {e}")
                    try:
                        driver.quit()
                    except Exception as quit_error:
                        self.logger.error(f"Erreur lors de la fermeture du WebDriver : {quit_error}")
                    finally:
                        self.current_size -= 1
                        self.logger.debug(f"Taille du pool décrémentée : {self.current_size}")


    def close_all(self):
        self.logger.info("Closing all WebDriver instances...")
        with self.lock:
            while not self.pool.empty():
                driver = self.pool.get()
                driver.quit()
            self.current_size = 0
        self.logger.info("All WebDriver instances closed.")


