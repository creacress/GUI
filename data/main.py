import sys
import logging
import os
from rpa_modules.debug import setup_logger
from rpa_modules.WebDriverPool import WebDriverPool
from rpa_modules.affranchigo import AffranchigoRPA
from rpa_modules.dematerialisation import CasDematerialisationRPA
from rpa_modules.extraction_odysse import ExtractionRPA
from rpa_modules.seres import SeresRPA

# Configuration du logger centralisé
logger = setup_logger('Affranchigo_premium_ROYE_PIC.log')

# Initialisation du pool de WebDrivers
DEFAULT_MAX_WORKERS = 5
DEFAULT_POOL_SIZE = 5
MAX_POOL_SIZE = 20
IDLE_TIMEOUT = 300

pool = WebDriverPool(initial_size=DEFAULT_POOL_SIZE, max_size=MAX_POOL_SIZE, idle_timeout=IDLE_TIMEOUT, logger=logger)

def main_rpa(rpa_name, max_workers=DEFAULT_MAX_WORKERS):
    """
    Point d'entrée principal pour gérer les différents RPA.
    """
    try:
        if rpa_name == "Affranchigo":
            affranchigo_rpa = AffranchigoRPA(pool, logger)
            affranchigo_rpa.main(max_workers=max_workers)

        elif rpa_name == "CasDematerialisation":
            demat_rpa = CasDematerialisationRPA(pool, logger)
            demat_rpa.main(max_workers=max_workers)

        elif rpa_name == "Extraction":
            extraction_rpa = ExtractionRPA(pool, logger)
            extraction_rpa.main(max_workers=max_workers)

        elif rpa_name == "Seres":
            seres_rpa = SeresRPA(pool, logger)
            seres_rpa.main(max_workers=max_workers)

        else:
            logger.error(f"RPA non reconnu: {rpa_name}")
            available_rpas = ["Affranchigo", "CasDematerialisation", "Extraction", "Seres"]
            logger.info(f"RPA disponibles : {', '.join(available_rpas)}")
            sys.exit(1)

    except Exception as e:
        logger.error(f"Erreur lors de l'exécution du RPA {rpa_name}: {e}", exc_info=True)
    finally:
        # Toujours fermer les WebDrivers à la fin
        logger.info("Fermeture de tous les WebDrivers...")
        pool.close_all()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        logger.error("Usage: python main.py <RPA_name> [<max_workers>]")
        sys.exit(1)

    # Récupère le nom du RPA à partir des arguments de la ligne de commande
    rpa_name = sys.argv[1]
    max_workers = int(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_MAX_WORKERS

    # Log du nom du RPA reçu
    logger.info(f"Nom du RPA reçu: {rpa_name}")
    logger.info(f"Nombre de threads maximum: {max_workers}")

    # Lancer le RPA correspondant
    main_rpa(rpa_name, max_workers)
