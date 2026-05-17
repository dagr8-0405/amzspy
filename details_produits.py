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
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)

def scraper_fiche_produit(driver, asin):
    url = f"https://www.amazon.fr/dp/{asin}"
    print(f"-> Scan profond et robuste de l'ASIN : {asin}...")
    
    try:
        driver.get(url)
        time.sleep(5) # Un peu plus de temps pour le scaling
        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # 1. Extraction Multi-Sélecteurs de la MARQUE
        brand = "Inconnu"
        brand_selectors = ["#bylineInfo", "#brand", "#amzn-byline-brand-", "a#bylineInfo"]
        for sel in brand_selectors:
            el = soup.select_one(sel)
            if el and el.text.strip():
                brand = el.text.strip()
                break
        brand = re.sub(r"(Visitez la boutique|Marque\s*:|Brand\s*:)", "", brand, flags=re.IGNORECASE).strip()
        
        # 2. Extraction Multi-Sélecteurs de la NOTE
        rating = "Pas de note"
        rating_selectors = ["span.a-icon-alt", "#acrPopover title", "i.a-icon-star span"]
        for sel in rating_selectors:
            el = soup.select_one(sel)
            if el:
                rating_match = re.search(r"([0-9][,.]?[0-9]?)", el.text)
                if rating_match:
                    rating = rating_match.group(1).replace(",", ".")
                    break
        
        # 3. Extraction Multi-Sélecteurs des AVIS
        reviews = "0"
        reviews_selectors = ["#acrCustomerReviewText", "span#acrCustomerReviewText", "a._cDE42_links_1b_7b"]
        for sel in reviews_selectors:
            el = soup.select_one(sel)
            if el and el.text.strip():
                reviews_raw = el.text
                reviews_clean = re.sub(r"[() \s,.]", "", reviews_raw)
                reviews_match = re.search(r"(\d+)", reviews_clean)
                if reviews_match:
                    reviews = reviews_match.group(1)
                    break
        
        # 4. Bullet Points (Sécurisés)
        bullets_list = [li.text.strip() for li in soup.select("#feature-bullets ul li span.a-list-item")]
        if not bullets_list:
            bullets_list = [span.text.strip() for span in soup.select("div#feature-bullets span.a-list-item")]
        while len(bullets_list) < 5:
            bullets_list.append("")
        
        # 5. Nombre d'images
        images = set([img['src'] for img in soup.select("#altImages img, #landingImage") if "src" in img.attrs and ("overlay" not in img['src'])])
        nb_images = len(images) if images else 1
        
        # 6. Description
        desc_el = soup.select_one("#productDescription, div#productDescription_feature_div")
        desc = desc_el.text.strip() if desc_el else "Aucune description"
        
        # 7. Caractéristiques techniques (Scraping du tableau ET des puces de détails)
        specs = {}
        for row in soup.select("#prodDetails table tr, #detailBullets_feature_div li, #productDetails_techSpec_section_1 tr"):
            th = row.select_one("th, span.a-text-bold, td.label")
            td = row.select_one("td, span:not(.a-text-bold), td.value")
            if th and td:
                k = th.text.replace(":", "").replace("›", "").strip()
                v = td.text.strip()
                if k and v and len(k) < 50: # Évite d'attraper des paragraphes entiers
                    specs[k] = v
        specs_str = str(specs) if specs else "Aucune caractéristique"
        
        # 8. Recherche Algorithmique du BSR (Indispensable pour le Scaling)
        bsr_text = "Non trouvé"
        page_text = soup.get_text()
        # Test de plusieurs expressions régulières (Regex) utilisées par Amazon France
        patterns = [
            r"Classement des meilleures ventes d'Amazon[\s\S]*?#([0-9\s,.]+) en",
            r"N°([0-9\s,.]+) dans High-Tech",
            r"N°([0-9\s,.]+) en",
            r"Classement des meilleures ventes[\s\S]*?#([0-9\s,.]+)"
        ]
        for pattern in patterns:
            bsr_match = re.search(pattern, page_text, re.IGNORECASE)
            if bsr_match:
                bsr_text = f"#{bsr_match.group(1).replace(' ', '').strip()}"
                break
        
        if bsr_text == "Non trouvé":
            for key, val in specs.items():
                if "Classement" in key or "Ventes" in key:
                    bsr_text = val
                    break

        # 9. Déclinaisons
        variants = [el.text.strip() for el in soup.select("#twister ul li span.a-button-text, div.inline-twister-row span.a-size-base")]
        nb_variants = len(variants)
        variants_str = ", ".join(variants) if variants else "Aucune déclinaison"

        return {
            "ASIN": asin,
            "Marque": brand,
            "Note": rating,
            "Nb_Avis": int(reviews) if reviews.isdigit() else 0,
            "Bullet_1": bullets_list[0],
            "Bullet_2": bullets_list[1],
            "Bullet_3": bullets_list[2],
            "Bullet_4": bullets_list[3],
            "Bullet_5": bullets_list[4],
            "Nombre_Images": nb_images,
            "Description": desc,
            "Caracteristiques": specs_str,
            "BSR_Categories": bsr_text,
            "Declinaisons": variants_str,
            "Nb_Declinaisons": nb_variants,
            "Date_Analyse": datetime.now().strftime("%Y-%m-%d")
        }
        
    except Exception as e:
        print(f"Erreur sur l'ASIN {asin} : {e}")
        return None

def executer_phase_2():
    chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)
    if not os.path.exists(chemin_fichier):
        print("Erreur : Fichier Excel introuvable.")
        return

    df_fiches = pd.read_excel(chemin_fichier, sheet_name="Fiches_Produits")
    df_clst = pd.read_excel(chemin_fichier, sheet_name="Suivi_Classement")
    
    # Nettoyage des colonnes obsolètes pour restructurer proprement
    if "Titre_Complet" in df_fiches.columns:
        df_fiches = df_fiches.drop(columns=["Titre_Complet"])
    if "Attributs" in df_fiches.columns:
        df_fiches = df_fiches.drop(columns=["Attributs"])

    # On cible les fiches vides ou jamais analysées
    if "Note" not in df_fiches.columns:
        asins_a_scanner = df_fiches["ASIN"].tolist()
    else:
        asins_a_scanner = df_fiches[(df_fiches["Marque"] == "À collecter en Phase 2") | (df_fiches["Marque"] == "Inconnu") | (df_fiches["Note"].isna())]["ASIN"].tolist()

    if not asins_a_scanner:
        print("Tous les produits sont parfaitement à jour. Aucun scan requis.")
        return

    # On monte le lot à 20 produits par session pour accélérer le scaling
    asins_a_scanner = asins_a_scanner[:20]
    print(f"Scaling : Analyse de {len(asins_a_scanner)} fiches produits avec l'algorithme de secours...")

    driver = configurer_driver()
    resultats = []
    
    try:
        for asin in asins_a_scanner:
            info = scraper_fiche_produit(driver, asin)
            if info:
                resultats.append(info)
            time.sleep(4)
    finally:
        driver.quit()

    if resultats:
        for res in resultats:
            asin = res["ASIN"]
            for col in res.keys():
                df_fiches.loc[df_fiches["ASIN"] == asin, col] = res[col]

        with pd.ExcelWriter(chemin_fichier, engine="openpyxl") as writer:
            df_clst.to_excel(writer, sheet_name="Suivi_Classement", index=False)
            df_fiches.to_excel(writer, sheet_name="Fiches_Produits", index=False)
            
        print("Fichier Excel optimisé et mis à jour !")
    else:
        print("Aucune donnée récoltée.")

if __name__ == "__main__":
    executer_phase_2()
