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
    # Simulation d'un vrai utilisateur français
    chrome_options.add_argument("--lang=fr-FR")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def scraper_fiche_produit(driver, asin):
    # On force la langue dans l'URL avec les paramètres Amazon
    url = f"https://www.amazon.fr/dp/{asin}?language=fr_FR&currency=EUR"
    print(f"-> Scan chirurgical de l'ASIN : {asin}")
    
    try:
        driver.get(url)
        time.sleep(6) # Pause pour laisser l'HTML s'injecter
        
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, "html.parser")
        
        # SÉCURITÉ ANTI-BOT : Si Amazon nous montre un captcha ou une page d'erreur
        if "captcha" in html_content.lower() or "api-services-support" in html_content.lower() or "sorry" in html_content.lower():
            print(f"⚠️ Alerte : Amazon a bloqué l'accès pour l'ASIN {asin}. Passage au suivant pour ne pas corrompre le fichier.")
            return None

        # 1. Extraction propre de la Marque
        brand = "Inconnu"
        brand_el = soup.select_one("#bylineInfo, #brand, #amzn-byline-brand-, a#bylineInfo")
        if brand_el:
            brand = re.sub(r"(Visitez la boutique|Marque\s*:|Brand\s*:|Visit the|Store)", "", brand_el.text, flags=re.IGNORECASE).strip()
        
        # 2. Note (Extraction brute du texte)
        rating_pure = 0.0
        for sel in ["span.a-icon-alt", "#acrPopover title", ".a-star-5", "i.a-icon-star"]:
            el = soup.select_one(sel)
            if el and el.text.strip():
                match = re.search(r"([0-9][,.]?[0-9]?)", el.text)
                if match:
                    rating_pure = float(match.group(1).replace(",", "."))
                    break
        
        # 3. Nombre d'avis (Nettoyage numérique radical)
        reviews_pure = 0
        reviews_el = soup.select_one("#acrCustomerReviewText, span#acrCustomerReviewText, #acrCustomerReviewLink")
        if reviews_el and reviews_el.text.strip():
            txt = reviews_el.text.replace(" ", "").replace(" ", "").replace(",", "").replace(".", "")
            match = re.search(r"(\d+)", txt)
            if match:
                reviews_pure = int(match.group(1))
        
        # 4. Bullet Points éclatés (Pas de colonne fusionnée)
        bullets = [li.text.strip() for li in soup.select("#feature-bullets ul li span.a-list-item")]
        if not bullets:
            bullets = [span.text.strip() for span in soup.select("div#feature-bullets span.a-list-item")]
        while len(bullets) < 5:
            bullets.append("")
        
        # 5. Extraction du tableau complet des caractéristiques
        specs = {}
        for row in soup.select("#prodDetails table tr, #detailBullets_feature_div li, #productDetails_techSpec_section_1 tr, .pdDetailsTable tr"):
            th = row.select_one("th, span.a-text-bold, td.label")
            td = row.select_one("td, span:not(.a-text-bold), td.value")
            if th and td:
                k = th.text.replace(":", "").replace("›", "").strip()
                v = td.text.strip()
                if k and v and len(k) < 60:
                    specs[k] = v
        specs_str = str(specs) if specs else "Aucune caractéristique"
        
        # 6. ISOLATION DU BSR (Peu importe la langue ou le format)
        bsr_text = "Non trouvé"
        # Recherche prioritaire dans le tableau technique qu'on vient de créer
        for k, v in specs.items():
            if any(x in k.lower() for x in ["classement", "ventes", "rank", "bestsellers"]):
                bsr_text = v.replace("\n", " ").strip()
                # Nettoyage si le texte est trop long
                bsr_text = re.sub(r"\s+", " ", bsr_text)
                break
        
        # Recherche de secours dans le texte brut complet
        if bsr_text == "Non trouvé":
            page_text = soup.get_text()
            patterns = [
                r"Classement des meilleures ventes d'Amazon[\s\S]*?#([0-9\s,.]+) (?:dans|en)",
                r"Amazon Bestsellers Rank[\s\S]*?#([0-9\s,.]+) (?:in|inside)",
                r"N°([0-9\s,.]+) (?:dans|en|in)"
            ]
            for pattern in patterns:
                match = re.search(pattern, page_text, re.IGNORECASE)
                if match:
                    bsr_text = f"#{match.group(1).replace(' ', '').strip()}"
                    break

        # 7. Images & Déclinaisons
        images = set([img['src'] for img in soup.select("#altImages img, #landingImage") if "src" in img.attrs and ("overlay" not in img['src'])])
        nb_images = len(images) if images else 1
        
        desc_el = soup.select_one("#productDescription, div#productDescription_feature_div")
        desc = desc_el.text.strip() if desc_el else "Aucune description"

        variants = [el.text.strip() for el in soup.select("#twister ul li span.a-button-text")]
        variants_str = ", ".join(variants) if variants else "Aucune déclinaison"

        return {
            "ASIN": asin,
            "Marque": brand,
            "Note": rating_pure,
            "Nb_Avis": reviews_pure,
            "Nombre_Images": nb_images,
            "Description": desc,
            "Caracteristiques": specs_str,
            "BSR_Categories": bsr_text,
            "Declinaisons": variants_str,
            "Nb_Declinaisons": len(variants),
            "Bullet_1": bullets[0],
            "Bullet_2": bullets[1],
            "Bullet_3": bullets[2],
            "Bullet_4": bullets[3],
            "Bullet_5": bullets[4],
            "Date_Analyse": datetime.now().strftime("%Y-%m-%d")
        }
    except Exception as e:
        print(f"Erreur technique sur l'ASIN {asin}: {e}")
        return None

def executer_phase_2():
    chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)
    if not os.path.exists(chemin_fichier):
        print("Erreur : Le fichier Excel central n'existe pas.")
        return

    df_fiches = pd.read_excel(chemin_fichier, sheet_name="Fiches_Produits")
    df_clst = pd.read_excel(chemin_fichier, sheet_name="Suivi_Classement")
    
    # Nettoyage structurel définitif des anciennes colonnes parasites
    for col in ["Titre_Complet", "Attributs", "Bullet_Points"]:
        if col in df_fiches.columns:
            df_fiches = df_fiches.drop(columns=[col])

    # Cibler uniquement les vraies fiches non remplies (évite d'écraser le propre avec du vide)
    if "Note" not in df_fiches.columns:
        asins_a_scanner = df_fiches["ASIN"].tolist()
    else:
        asins_a_scanner = df_fiches[
            (df_fiches["Marque"] == "À collecter en Phase 2") | 
            (df_fiches["Marque"] == "Inconnu") |
            (df_fiches["BSR_Categories"] == "Non trouvé") |
            (df_fiches["BSR_Categories"].isna())
        ]["ASIN"].tolist()

    if not asins_a_scanner:
        print("Toutes les fiches du fichier sont parfaitement propres et valides.")
        return

    # On traite un petit lot de 10 lignes pour valider la qualité du nettoyage
    asins_a_scanner = asins_a_scanner[:10]
    print(f"Lancement de la phase de nettoyage sur {len(asins_a_scanner)} produits...")

    driver = configurer_driver()
    resultats = []
    
    try:
        for asin in asins_a_scanner:
            info = scraper_fiche_produit(driver, asin)
            if info: # Si info est None (blocage Amazon), on n'écrase rien !
                resultats.append(info)
            time.sleep(5)
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
        print("Fichier central mis à jour avec succès.")
    else:
        print("Aucune donnée n'a été modifiée lors de cette session (protection anti-blocage).")

if __name__ == "__main__":
    executer_phase_2()
