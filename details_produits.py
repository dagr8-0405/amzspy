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
    chrome_options.add_experimental_option('prefs', {'intl.accept_languages': 'fr,fr-FR'})
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=chrome_options)

def scraper_fiche_produit(driver, asin):
    url = f"https://www.amazon.fr/dp/{asin}"
    print(f"-> Scan et structuration de l'ASIN : {asin}")
    
    try:
        driver.get(url)
        time.sleep(5)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # 1. Extraction & Nettoyage de la Marque
        brand = "Inconnu"
        brand_el = soup.select_one("#bylineInfo, #brand, #amzn-byline-brand-")
        if brand_el:
            brand = re.sub(r"(Visitez la boutique|Marque\s*:|Brand\s*:|Visit the|Store)", "", brand_el.text, flags=re.IGNORECASE).strip()
        
        # 2. Extraction & Nettoyage de la Note (Nombre pur)
        rating_pure = None
        rating_el = soup.select_one("span.a-icon-alt, #acrPopover title, .a-star-5")
        if rating_el:
            # Recherche d'un chiffre comme 4.5 ou 4,5 ou 4.5 out of 5
            rating_match = re.search(r"([0-9][,.]?[0-9]?)", rating_el.text)
            if rating_match:
                rating_pure = float(rating_match.group(1).replace(",", "."))
        
        # 3. Extraction & Nettoyage du Nombre d'avis (Chiffre entier pur)
        reviews_pure = 0
        reviews_el = soup.select_one("#acrCustomerReviewText, span#acrCustomerReviewText, #acrCustomerReviewLink")
        if reviews_pure == 0 and reviews_el:
            # Enlever agressivement les parenthèses, espaces insécables, virgules et points
            reviews_clean = re.sub(r"[() \s,.]", "", reviews_el.text)
            reviews_match = re.search(r"(\d+)", reviews_clean)
            if reviews_match:
                reviews_pure = int(reviews_match.group(1))
        
        # 4. Extraction dispatchée des Bullet Points (Pas de colonne globale)
        bullets = [li.text.strip() for li in soup.select("#feature-bullets ul li span.a-list-item")]
        if not bullets:
            bullets = [span.text.strip() for span in soup.select("div#feature-bullets span.a-list-item")]
        while len(bullets) < 5:
            bullets.append("")
        
        # 5. Nombre d'images
        images = set([img['src'] for img in soup.select("#altImages img, #landingImage") if "src" in img.attrs and ("overlay" not in img['src'])])
        nb_images = len(images) if images else 1
        
        # 6. Description
        desc_el = soup.select_one("#productDescription")
        desc = desc_el.text.strip() if desc_el else "Aucune description"
        
        # 7. Lecture et isolation du Tableau des Caractéristiques
        specs = {}
        for row in soup.select("#prodDetails table tr, #detailBullets_feature_div li, #productDetails_techSpec_section_1 tr"):
            th = row.select_one("th, span.a-text-bold, td.label")
            td = row.select_one("td, span:not(.a-text-bold), td.value")
            if th and td:
                k = th.text.replace(":", "").replace("›", "").strip()
                v = td.text.strip()
                if k and v and len(k) < 50:
                    specs[k] = v
        specs_str = str(specs) if specs else "Aucune caractéristique"
        
        # 8. EXTRACTION INTELLIGENTE DU BSR DEPUIS LES CARACTÉRISTIQUES (Correction Scaling)
        bsr_text = "Non trouvé"
        # On fouille d'abord dans le dictionnaire technique généré juste au-dessus
        for key, val in specs.items():
            if any(x in key.lower() for x in ["classement", "ventes", "rank", "bestsellers"]):
                bsr_text = val.replace("\n", " ").strip()
                break
        
        # Méthode de secours B : Si pas dans le tableau, on fouille le texte brut de la page
        if bsr_text == "Non trouvé":
            page_text = soup.get_text()
            patterns = [
                r"Classement des meilleures ventes d'Amazon[\s\S]*?#([0-9\s,.]+) en",
                r"Amazon Bestsellers Rank[\s\S]*?#([0-9\s,.]+) in",
                r"N°([0-9\s,.]+) (?:dans|en|in)",
                r"Rank[\s\S]*?#([0-9\s,.]+)"
            ]
            for pattern in patterns:
                bsr_match = re.search(pattern, page_text, re.IGNORECASE)
                if bsr_match:
                    bsr_text = f"#{bsr_match.group(1).replace(' ', '').strip()}"
                    break

        # 9. Déclinaisons
        variants = [el.text.strip() for el in soup.select("#twister ul li span.a-button-text")]
        nb_variants = len(variants)
        variants_str = ", ".join(variants) if variants else "Aucune déclinaison"

        return {
            "ASIN": asin,
            "Marque": brand,
            "Note": rating_pure if rating_pure is not None else 0.0,
            "Nb_Avis": reviews_pure,
            "Nombre_Images": nb_images,
            "Description": desc,
            "Caracteristiques": specs_str,
            "BSR_Categories": bsr_text,
            "Declinaisons": variants_str,
            "Nb_Declinaisons": nb_variants,
            "Bullet_1": bullets[0],
            "Bullet_2": bullets[1],
            "Bullet_3": bullets[2],
            "Bullet_4": bullets[3],
            "Bullet_5": bullets[4],
            "Date_Analyse": datetime.now().strftime("%Y-%m-%d")
        }
    except Exception as e:
        print(f"Erreur ASIN {asin}: {e}")
        return None

def executer_phase_2():
    chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)
    if not os.path.exists(chemin_fichier):
        print("Fichier Excel central introuvable.")
        return

    df_fiches = pd.read_excel(chemin_fichier, sheet_name="Fiches_Produits")
    df_clst = pd.read_excel(chemin_fichier, sheet_name="Suivi_Classement")
    
    # Nettoyage strict des colonnes obsolètes pour éviter les doublons structurels
    colonnes_a_supprimer = ["Titre_Complet", "Attributs", "Bullet_Points"]
    for col in colonnes_a_supprimer:
        if col in df_fiches.columns:
            df_fiches = df_fiches.drop(columns=[col])

    # Détection des lignes qui ont besoin d'être scannées ou corrigées
    if "Note" not in df_fiches.columns:
        asins_a_scanner = df_fiches["ASIN"].tolist()
    else:
        # On rescane si la marque est inconnue, si la note est à 0, ou si le BSR a échoué (Non trouvé)
        asins_a_scanner = df_fiches[
            (df_fiches["Marque"] == "À collecter en Phase 2") | 
            (df_fiches["Marque"] == "Inconnu") | 
            (df_fiches["BSR_Categories"] == "Non trouvé") |
            (df_fiches["Note"] == 0.0)
        ]["ASIN"].tolist()

    if not asins_a_scanner:
        print("Toutes les fiches produits sont parfaitement structurées pour le scaling !")
        return

    # Extraction par lot de 15 pour la sécurité du compte GitHub
    asins_a_scanner = asins_a_scanner[:15]
    print(f"Structure & Scaling : Analyse approfondie de {len(asins_a_scanner)} fiches...")

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
            # Mise à jour ou création dynamique des colonnes propres
            for col in res.keys():
                df_fiches.loc[df_fiches["ASIN"] == asin, col] = res[col]

        # Réécriture propre
        with pd.ExcelWriter(chemin_fichier, engine="openpyxl") as writer:
            df_clst.to_excel(writer, sheet_name="Suivi_Classement", index=False)
            df_fiches.to_excel(writer, sheet_name="Fiches_Produits", index=False)
        print("Sauvegarde de la nouvelle structure réussie.")
    else:
        print("Aucun changement appliqué.")

if __name__ == "__main__":
    executer_phase_2()
