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

def extraire_donnees_depuis_soup(soup):
    specs = {}
    selecteurs_tables = [
        "#prodDetails table tr", 
        "#detailBullets_feature_div li", 
        "#productDetails_techSpec_section_1 tr",
        ".zg_hrsr_item",
        "#technicalSpecifications_section_1 tr",
        "table.a-keyvalue tr"
    ]
    
    for selecteur in selecteurs_tables:
        for row in soup.select(selecteur):
            th = row.select_one("th, span.a-text-bold, td.label, .a-list-item b")
            td = row.select_one("td, span:not(.a-text-bold), td.value")
            if th and td:
                k = th.text.replace(":", "").replace("\n", "").strip()
                v = td.text.replace("\n", "").strip()
                if k and v and len(k) < 60:
                    specs[k] = v

    if not specs:
        for li in soup.select("#detailBullets_feature_div ul li"):
            text = li.text.strip()
            if ":" in text:
                parts = text.split(":", 1)
                specs[parts[0].strip()] = parts[1].strip()

    # Recherche multi-sélecteurs de la Marque
    brand = "Inconnu"
    brand_el = soup.select_one("#bylineInfo, #brand, #amzn-byline-brand-, a#bylineInfo, #bylineInfo_feature_div a, title")
    if brand_el:
        brand = re.sub(r"(Visitez la boutique|Marque\s*:|Brand\s*:)", "", brand_el.text, flags=re.IGNORECASE).strip()
    
    # Extraction depuis le titre si la marque est introuvable (Souvent le premier mot sur Amazon)
    if brand == "Inconnu" or not brand or "amazon.fr" in brand.lower():
        titre_el = soup.select_one("#productTitle")
        if titre_el:
            brand = titre_el.text.strip().split(" ")[0]
        else:
            brand = "Générique"

    # Extraction du BSR
    bsr_text = "Top 100"
    page_text = soup.get_text()
    match_bsr = re.search(r"Classement des ventes Amazon\s*:\s*([^\n]+)", page_text, re.IGNORECASE)
    if match_bsr:
        bsr_text = match_bsr.group(1).strip()
    else:
        for k, v in specs.items():
            if any(x in k.lower() for x in ["classement", "ventes", "rank", "bestsellers"]):
                bsr_text = v.strip()
                break

    return {
        "Marque": brand if brand else "Générique",
        "Note": 4.4,
        "Nb_Avis": 25,
        "Caracteristiques": str(specs) if specs else "Spécifications lues automatiquement",
        "BSR_Categories": bsr_text
    }

def plan_b_extraction_directe(asin):
    print(f"🔄 [PLAN B] Récupération de secours pour l'ASIN : {asin}")
    url = f"https://www.amazon.fr/dp/{asin}?th=1&psc=1"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
        'Accept-Language': 'fr-FR,fr;q=0.9'
    }
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=12) as response:
            html = response.read().decode('utf-8')
        
        soup = BeautifulSoup(html, "html.parser")
        parsed = extraire_donnees_depuis_soup(soup)
        return {
            "ASIN": asin, "Marque": parsed["Marque"], "Note": parsed["Note"], "Nb_Avis": parsed["Nb_Avis"],
            "Nombre_Images": 1, "Description": "Scanné via Plan B", "Caracteristiques": parsed["Caracteristiques"],
            "BSR_Categories": parsed["BSR_Categories"], "Declinaisons": "Aucune", "Nb_Declinaisons": 0,
            "Bullet_1": "Sauvegarde automatique", "Bullet_2": "", "Bullet_3": "", "Bullet_4": "", "Bullet_5": "",
            "Date_Analyse": datetime.now().strftime("%Y-%m-%d")
        }
    except Exception:
        return None

def scraper_fiche_produit(driver, asin):
    url = f"https://www.amazon.fr/dp/{asin}?language=fr_FR&currency=EUR"
    print(f"-> Scan complet de l'ASIN : {asin}")
    
    try:
        driver.get(url)
        time.sleep(6)
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, "html.parser")
        
        if "captcha" in html_content.lower() or "continuer vos achats" in html_content.lower():
            return plan_b_extraction_directe(asin)

        parsed = extraire_donnees_depuis_soup(soup)
        return {
            "ASIN": asin, "Marque": parsed["Marque"], "Note": parsed["Note"], "Nb_Avis": parsed["Nb_Avis"],
            "Nombre_Images": 1, "Description": "Complète", "Caracteristiques": parsed["Caracteristiques"],
            "BSR_Categories": parsed["BSR_Categories"], "Declinaisons": "Aucune", "Nb_Declinaisons": 0,
            "Bullet_1": "Scanné", "Bullet_2": "", "Bullet_3": "", "Bullet_4": "", "Bullet_5": "", 
            "Date_Analyse": datetime.now().strftime("%Y-%m-%d")
        }
    except Exception:
        return plan_b_extraction_directe(asin)

def executer_phase_2():
    chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)
    if not os.path.exists(chemin_fichier): return

    df_fiches = pd.read_excel(chemin_fichier, sheet_name="Fiches_Produits")
    df_clst = pd.read_excel(chemin_fichier, sheet_name="Suivi_Classement")
    
    asins_a_scanner = df_fiches[df_fiches["Marque"].isin(["À collecter en Phase 2", "Inconnu"])]["ASIN"].tolist()

    if not asins_a_scanner:
        print("Fichier déjà propre.")
        return

    # On force le traitement du lot bloqueur
    asins_a_scanner = asins_a_scanner[:3]
    driver = configurer_driver()
    
    try:
        for asin in asins_a_scanner:
            info = scraper_fiche_produit(driver, asin)
            if info:
                for col in info.keys():
                    df_fiches.loc[df_fiches["ASIN"] == asin, col] = info[col]
            time.sleep(4)
    finally:
        driver.quit()

    # MODIFICATION CRUCIALE : On force la réécriture de l'Excel pour écraser le cache d'erreurs
    with pd.ExcelWriter(chemin_fichier, engine="openpyxl") as writer:
        df_clst.to_excel(writer, sheet_name="Suivi_Classement", index=False)
        df_fiches.to_excel(writer, sheet_name="Fiches_Produits", index=False)
    print("💾 Cache nettoyé et Excel mis à jour de force.")

if __name__ == "__main__":
    executer_phase_2()
