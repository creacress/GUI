import os
import time
import threading
from queue import Queue
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from rpa_modules.debug import setup_logger

class WebDriverPool:
    def __init__(self, initial_size=5, max_size=30, idle_timeout=100, logger=None):
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
        self.logger = logger or setup_logger('WebDriverPool.log')
        self.logger.debug("WebDriverPool initialized with max_size=%d", max_size)

        # Pré-charger les instances initiales de WebDriver
        for _ in range(initial_size):
            driver = self.create_driver()
            self.pool.put(driver)
            self.current_size += 1

    def create_driver(self):
        """
        Crée une nouvelle instance de WebDriver configurée avec les options requises.
        """
        try:
            self.logger.debug("Creating a new WebDriver instance")

            driver_path = r'data/driver/msedgedriver.exe'

            service = Service(driver_path)
            options = Options()
            options.add_argument("--headless")
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

            # Naviguer vers l'URL de départ
            driver.get("https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/index/acces-dc?profil=ADV")
            driver.last_used_time = time.time()

            self.logger.debug("WebDriver instance created successfully")
            return driver
        except Exception as e:
            self.logger.error(f"Erreur lors de la création du WebDriver: {e}")
            raise

    def get_driver(self, url=None):
        """
        Obtient un WebDriver du pool, en créant un nouveau si nécessaire.
        """
        with self.lock:
            if not self.pool.empty():
                driver = self.pool.get()
                self.logger.debug("Réutilisation d'une instance WebDriver existante.")
                try:
                    driver.execute_script("return document.readyState")
                except Exception as e:
                    self.logger.warning(f"WebDriver inactif, recréation: {e}")
                    driver = self.create_driver()
            else:
                if self.current_size < self.max_size:
                    driver = self.create_driver()
                    self.current_size += 1
                    self.logger.debug(f"Nouvelle instance créée (taille actuelle: {self.current_size}).")
                else:
                    raise Exception("Taille maximale du pool atteinte.")

            if url:
                driver.get(url)
            driver.last_used_time = time.time()
            return driver

    def return_driver(self, driver):
        with self.lock:
            if driver:
                try:
                    driver.get("https://www.deviscontrat.net-courrier.extra.laposte.fr/appli/ihm/index/acces-dc?profil=ADV")
                    status = driver.execute_script("return document.readyState")
                    if status != "complete":
                        raise Exception("WebDriver non opérationnel.")

                    self.pool.put(driver)
                    self.logger.debug(f"WebDriver remis dans le pool. Taille actuelle: {self.pool.qsize()}.")
                except Exception as e:
                    self.logger.warning(f"Driver inactif, suppression: {e}")
                    driver.quit()
                    self.current_size -= 1
                    self.logger.debug(f"Taille du pool décrémentée : {self.current_size}")
                    # Recréer un driver pour maintenir la taille du pool
                    if self.current_size < self.max_size:
                        new_driver = self.create_driver()
                        self.pool.put(new_driver)
                        self.current_size += 1
                        self.logger.debug(f"Nouvelle instance créée après suppression d'une défectueuse.")


    def close_all(self):
        """
        Ferme toutes les instances de WebDrivers.
        """
        self.logger.info("Fermeture de toutes les instances WebDriver...")
        with self.lock:
            while not self.pool.empty():
                driver = self.pool.get()
                driver.quit()
            self.current_size = 0
        self.logger.info("Toutes les instances WebDriver ont été fermées.")
