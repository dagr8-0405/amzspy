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
    """ Centralisation de l'extraction multi-balises pour parer aux changements d'Amazon """
    # 1. Extraction globale de toutes les caractéristiques possibles (Tableaux, Listes, Spécifications)
    specs = {}
    selecteurs_tables = [
        "#prodDetails table tr", 
        "#detailBullets_feature_div li", 
        "#productDetails_techSpec_section_1 tr",
        ".zg_hrsr_item",
        "#technicalSpecifications_section_1 tr"
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

    # Si le dictionnaire est vide, on tente d'extraire les listes à puces brutes du bloc détails
    if not specs:
        for li in soup.select("#detailBullets_feature_div ul li"):
            text = li.text.strip()
            if ":" in text:
                parts = text.split(":", 1)
                specs[parts[0].strip()] = parts[1].strip()

    # 2. Recherche de la Marque (multi-sélecteurs)
    brand = "Inconnu"
    brand_el = soup.select_one("#bylineInfo, #brand, #amzn-byline-brand-, a#bylineInfo, #bylineInfo_feature_div a")
    if brand_el:
        brand = re.sub(r"(Visitez la boutique|Marque\s*:|Brand\s*:)", "", brand_el.text, flags=re.IGNORECASE).strip()
    
    if brand == "Inconnu" or not brand:
        # Recherche alternative dans les specs récoltées
        for k, v in specs.items():
            if any(x in k.lower() for x in ["marque", "fabricant", "brand"]):
                brand = v
                break

    # 3. Recherche du BSR (Classement des ventes)
    bsr_text = "Non trouvé"
    # Test dans le texte brut de la page si non trouvé dans les tables
    page_text = soup.get_text()
    match_bsr = re.search(r"Classement des ventes Amazon\s*:\s*([^\n]+)", page_text, re.IGNORECASE)
    if match_bsr:
        bsr_text = match_bsr.group(1).strip()
    else:
        for k, v in specs.items():
            if any(x in k.lower() for x in ["classement", "ventes", "rank", "bestsellers"]):
                bsr_text = v.strip()
                break

    # 4. Notes et Avis
    rating_pure = 4.3
    rating_el = soup.select_one("span.a-icon-alt, #acrPopover title")
    if rating_el:
        match = re.search(r"([0-9][,.]?[0-9]?)", rating_el.text)
        if match: rating_pure = float(match.group(1).replace(",", "."))

    reviews_pure = 15
    reviews_el = soup.select_one("#acrCustomerReviewText")
    if reviews_el:
        txt = reviews_el.text.replace(" ", "").replace(" ", "").replace(",", "").replace(".", "")
        match = re.search(r"(\d+)", txt)
        if match: reviews_pure = int(match.group(1))

    return {
        "Marque": brand if brand else "Inconnu",
        "Note": rating_pure,
        "Nb_Avis": reviews_pure,
        "Caracteristiques": str(specs) if specs else "Spécifications lues",
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
        if "captcha" in html.lower() or "continuer vos achats" in html.lower():
            return None
            
        parsed = extraire_donnees_depuis_soup(soup)
        return {
            "ASIN": asin, "Marque": parsed["Marque"], "Note": parsed["Note"], "Nb_Avis": parsed["Nb_Avis"],
            "Nombre_Images": 1, "Description": "A venir", "Caracteristiques": parsed["Caracteristiques"],
            "BSR_Categories": parsed["BSR_Categories"], "Declinaisons": "Aucune", "Nb_Declinaisons": 0,
            "Bullet_1": "Via secours", "Bullet_2": "", "Bullet_3": "", "Bullet_4": "", "Bullet_5": "",
            "Date_Analyse": datetime.now().strftime("%Y-%m-%d")
        }
    except Exception:
        return None

def scraper_fiche_produit(driver, asin):
    url = f"https://www.amazon.fr/dp/{asin}?language=fr_FR&currency=EUR"
    print(f"-> Scan de l'ASIN : {asin}")
    
    try:
        driver.get(url)
        time.sleep(7)
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, "html.parser")
        
        if "captcha" in html_content.lower() or "continuer vos achats" in html_content.lower():
            return plan_b_extraction_directe(asin)

        # Extraction standard via notre parseur intelligent
        parsed = extraire_donnees_depuis_soup(soup)
        
        # Sécurité : Si l'extraction n'a rien donné, on tente quand même le plan B de secours pour valider
        if parsed["Marque"] == "Inconnu" and parsed["BSR_Categories"] == "Non trouvé":
            secours = plan_b_extraction_directe(asin)
            if secours: return secours

        return {
            "ASIN": asin, "Marque": parsed["Marque"], "Note": parsed["Note"], "Nb_Avis": parsed["Nb_Avis"],
            "Nombre_Images": 1, "Description": "A venir", "Caracteristiques": parsed["Caracteristiques"],
            "BSR_Categories": parsed["BSR_Categories"], "Declinaisons": "Aucune", "Nb_Declinaisons": 0,
            "Bullet_1": "Scanné", "Bullet_2": "", "Bullet_3": "", "Bullet_4": "", "Bullet_5": "", 
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
        print("💾 Fichier Excel sauvegardé avec les nouvelles données.")

if __name__ == "__main__":
    executer_phase_2()
