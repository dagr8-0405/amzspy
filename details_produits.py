import os
import time
import re
from datetime import datetime
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

NOM_FICHIER_CENTRAL = "historique_bestsellers.xlsx"

def configurer_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--lang=fr-FR")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)

def scraper_fiche_produit(driver, asin):
    url = f"https://www.amazon.fr/dp/{asin}?language=fr_FR&currency=EUR"
    print(f"-> Scan chirurgical de l'ASIN : {asin}")
    
    try:
        driver.get(url)
        time.sleep(6)
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, "html.parser")
        
        # FORCE L'AUTOCORRECTION EN CAS DE BLOCAGE
        if "captcha" in html_content.lower() or "api-services-support" in html_content.lower() or "sorry" in html_content.lower() or "robot" in html_content.lower():
            print(f"⚠️ [IA TRIGGER] Amazon a bloqué l'ASIN {asin}. Génération forcée du ticket d'anomalie.")
            with open("alerte_ia.txt", "w", encoding="utf-8") as f:
                f.write(f"ASIN CIBLE : {asin}\n")
                f.write(f"ANOMALIE DETECTEE : Blocage Securite Amazon (Captcha / Anti-Bot)\n")
                f.write(f"TEXTE DE LA PAGE DE BLOCAGE : {soup.get_text()[:1500]}\n")
            return None

        # Extraction classique (si ça passe)
        brand = "Inconnu"
        brand_el = soup.select_one("#bylineInfo, #brand, #amzn-byline-brand-")
        if brand_el:
            brand = re.sub(r"(Visitez la boutique|Marque\s*:|Brand\s*:)", "", brand_el.text, flags=re.IGNORECASE).strip()
        
        rating_pure = 0.0
        rating_el = soup.select_one("span.a-icon-alt, #acrPopover title")
        if rating_el:
            match = re.search(r"([0-9][,.]?[0-9]?)", rating_el.text)
            if match: rating_pure = float(match.group(1).replace(",", "."))
        
        reviews_pure = 0
        reviews_el = soup.select_one("#acrCustomerReviewText")
        if reviews_el:
            txt = reviews_el.text.replace(" ", "").replace(" ", "").replace(",", "").replace(".", "")
            match = re.search(r"(\d+)", txt)
            if match: reviews_pure = int(match.group(1))

        bullets = [li.text.strip() for li in soup.select("#feature-bullets ul li span.a-list-item")][:5]
        while len(bullets) < 5: bullets.append("")

        specs = {}
        for row in soup.select("#prodDetails table tr, #detailBullets_feature_div li"):
            th = row.select_one("th, span.a-text-bold")
            td = row.select_one("td, span:not(.a-text-bold)")
            if th and td:
                k = th.text.replace(":", "").strip()
                v = td.text.strip()
                if k and v and len(k) < 50: specs[k] = v
        specs_str = str(specs) if specs else "Aucune caractéristique"

        bsr_text = "Non trouvé"
        for k, v in specs.items():
            if any(x in k.lower() for x in ["classement", "ventes", "rank"]):
                bsr_text = v.replace("\n", " ").strip()
                break

        return {
            "ASIN": asin, "Marque": brand, "Note": rating_pure, "Nb_Avis": reviews_pure,
            "Nombre_Images": 1, "Description": "A venir", "Caracteristiques": specs_str,
            "BSR_Categories": bsr_text, "Declinaisons": "Aucune", "Nb_Declinaisons": 0,
            "Bullet_1": bullets[0], "Bullet_2": bullets[1], "Bullet_3": bullets[2],
            "Bullet_4": bullets[3], "Bullet_5": bullets[4], "Date_Analyse": datetime.now().strftime("%Y-%m-%d")
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
        print("Tout est OK.")
        return

    # On teste sur un micro-lot de 3 produits pour forcer la detection
    asins_a_scanner = asins_a_scanner[:3]
    driver = configurer_driver()
    
    try:
        for asin in asins_a_scanner:
            info = scraper_fiche_produit(driver, asin)
            if info:
                for col in info.keys():
                    df_fiches.loc[df_fiches["ASIN"] == asin, col] = info[col]
            time.sleep(5)
    finally:
        driver.quit()

    with pd.ExcelWriter(chemin_fichier, engine="openpyxl") as writer:
        df_clst.to_excel(writer, sheet_name="Suivi_Classement", index=False)
        df_fiches.to_excel(writer, sheet_name="Fiches_Produits", index=False)

if __name__ == "__main__":
    executer_phase_2()
