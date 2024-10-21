import logging
from logging.handlers import RotatingFileHandler
import datetime
from selenium import webdriver
import os

def setup_logger(log_file='app.log', level=logging.DEBUG):
    logger = logging.getLogger('central_logger')

    if not logger.hasHandlers():
        logger.setLevel(level)

        # Créer le dossier des logs si nécessaire
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir)
            print(f"Dossier {log_dir} créé.")

        # Créer le fichier de log si nécessaire
        if not os.path.exists(log_file):
            open(log_file, 'w').close()
            print(f"Fichier {log_file} créé.")

        # Création d'un gestionnaire de fichier avec rotation
        file_handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=5, encoding='utf-8')
        file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)

        # Création d'un gestionnaire pour afficher les logs dans la console (optionnel)
        console_handler = logging.StreamHandler()
        console_formatter = logging.Formatter('%(name)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(console_formatter)
        logger.addHandler(console_handler)

    return logger

