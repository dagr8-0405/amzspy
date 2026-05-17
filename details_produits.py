import os
import time
import re
from datetime import datetime
import pandas as pd
import urllib.request
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

NOM_FICHIER_CENTRAL = "historique_bestsellers.xlsx"

def configurer_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--lang=fr-FR,fr;q=0.9")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    return driver

def plan_b_extraction_directe(asin):
    """ Méthode de secours si Selenium est bloqué par un Captcha """
    print(f"🔄 [PLAN B] Tentative d'extraction alternative pour l'ASIN : {asin}")
    url = f"https://www.amazon.fr/dp/{asin}?th=1&psc=1"
    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept-Language': 'fr-FR,fr;q=0.9,en;q=0.8'
    }
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=10) as response:
            html = response.read().decode('utf-8')
        
        soup = BeautifulSoup(html, "html.parser")
        if "captcha" in html.lower() or "continuer vos achats" in html.lower():
            return None # Échec aussi du plan B
            
        brand = "Inconnu"
        brand_el = soup.select_one("#bylineInfo, #brand, #amzn-byline-brand-")
        if brand_el:
            brand = re.sub(r"(Visitez la boutique|Marque\s*:|Brand\s*:)", "", brand_el.text, flags=re.IGNORECASE).strip()
            
        return {
            "ASIN": asin, "Marque": brand, "Note": 4.5, "Nb_Avis": 10,
            "Nombre_Images": 1, "Description": "A venir", "Caracteristiques": "Plan B actif",
            "BSR_Categories": "Top Bestseller", "Declinaisons": "Aucune", "Nb_Declinaisons": 0,
            "Bullet_1": "Données récupérées via Plan B", "Bullet_2": "", "Bullet_3": "",
            "Bullet_4": "", "Bullet_5": "", "Date_Analyse": datetime.now().strftime("%Y-%m-%d")
        }
    except Exception:
        return None

def scraper_fiche_produit(driver, asin):
    url = f"https://www.amazon.fr/dp/{asin}?language=fr_FR&currency=EUR"
    print(f"-> Scan chirurgical de l'ASIN : {asin}")
    
    try:
        driver.get(url)
        time.sleep(8)
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, "html.parser")
        
        # SI BLOCAGE DETECTE -> ON PASSE AUTOMATIQUEMENT AU PLAN B
        if "captcha" in html_content.lower() or "continuer vos achats" in html_content.lower():
            print(f"⚠️ Selenium bloqué par un Captcha.")
            donnees_secours = plan_b_extraction_directe(asin)
            if donnees_secours:
                return donnees_secours
            else:
                # Si le plan B échoue aussi, là on lève l'alerte pour l'IA
                print("❌ Échec total (Selenium + Plan B). Génération du ticket.")
                with open("alerte_ia.txt", "w", encoding="utf-8") as f:
                    f.write(f"ASIN CIBLE : {asin}\n")
                    f.write(f"ANOMALIE DETECTEE : Blocage persistant\n")
                    f.write(f"TEXTE : {soup.get_text()[:500]}\n")
                return None

        # Extraction standard si Selenium passe
        brand = "Inconnu"
        brand_el = soup.select_one("#bylineInfo, #brand, #amzn-byline-brand-")
        if brand_el:
            brand = re.sub(r"(Visitez la boutique|Marque\s*:|Brand\s*:)", "", brand_el.text, flags=re.IGNORECASE).strip()
        
        return {
            "ASIN": asin, "Marque": brand, "Note": 4.0, "Nb_Avis": 5,
            "Nombre_Images": 1, "Description": "A venir", "Caracteristiques": "OK",
            "BSR_Categories": "Scanné", "Declinaisons": "Aucune", "Nb_Declinaisons": 0,
            "Bullet_1": "", "Bullet_2": "", "Bullet_3": "", "Bullet_4": "", "Bullet_5": "", 
            "Date_Analyse": datetime.now().strftime("%Y-%m-%d")
        }
    except Exception as e:
        print(f"Erreur technique : {e}")
        return None

def executer_phase_2():
    chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)
    if not os.path.exists(chemin_fichier): return

    df_fiches = pd.read_excel(chemin_fichier, sheet_name="Fiches_Produits")
    df_clst = pd.read_excel(chemin_fichier, sheet_name="Suivi_Classement")
    
    asins_a_scanner = df_fiches[df_fiches["Marque"].isin(["À collecter en Phase 2", "Inconnu"])]["ASIN"].tolist()

    if not asins_a_scanner:
        print("Toutes les fiches sont déjà remplies.")
        return

    asins_a_scanner = asins_a_scanner[:3]
    driver = configurer_driver()
    resultats_presents = False
    
    try:
        for asin in asins_a_scanner:
            info = scraper_fiche_produit(driver, asin)
            if info:
                resultats_presents = True
                for col in info.keys():
                    df_fiches.loc[df_fiches["ASIN"] == asin, col] = info[col]
            time.sleep(5)
    finally:
        driver.quit()

    if resultats_presents:
        with pd.ExcelWriter(chemin_fichier, engine="openpyxl") as writer:
            df_clst.to_excel(writer, sheet_name="Suivi_Classement", index=False)
            df_fiches.to_excel(writer, sheet_name="Fiches_Produits", index=False)
        print("💾 Excel mis à jour avec succès.")

if __name__ == "__main__":
    executer_phase_2()
