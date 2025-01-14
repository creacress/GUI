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
logger = setup_logger('Affranchigo_ROYE.log')

# Initialisation du pool de WebDrivers
DEFAULT_MAX_WORKERS = 5
pool = WebDriverPool(initial_size=DEFAULT_MAX_WORKERS, max_size=30, idle_timeout=100, logger=None)

def main_rpa(rpa_name, max_workers=DEFAULT_MAX_WORKERS):
    """
    Point d'entrée principal pour gérer les différents RPA.
    """
    try:
        if rpa_name == "Affranchigo":
            affranchigo_rpa = AffranchigoRPA(pool, logger)
            affranchigo_rpa.main(max_workers=max_workers)  # Appel de la méthode principale

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
        logger.error(f"Erreur lors de l'exécution du RPA {rpa_name}: {e}")
    finally:
        # Toujours fermer les WebDrivers à la fin
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

    # Lancer le RPA correspondant
    main_rpa(rpa_name, max_workers)