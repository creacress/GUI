from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

driver_path = r'data/driver/msedgedriver.exe'

# Initialisation du WebDriver pour Edge
service = Service(driver_path)
options = Options()
driver = webdriver.Edge(service=service, options=options)

try:
    driver.get("https://www.google.com")
    
    # Attendre que l'élément de recherche soit visible et cliquable
    wait = WebDriverWait(driver, 10)
    button_delete = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#W0wltc > div")))
    button_delete.click()
    
    search_box = wait.until(EC.element_to_be_clickable((By.NAME, "q")))
    
    # Interagir avec l'élément de manière sécurisée
    search_box.send_keys("Test WebDriver")
    search_box.submit()

    print("Le test a fonctionné correctement.")
    
except Exception as e:
    print(f"Erreur durant le test : {e}")
finally:
    driver.quit()
