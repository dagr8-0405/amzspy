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
    print(f"-> Scan propre du produit {asin}...")
    
    try:
        driver.get(url)
        time.sleep(4)
        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # 1. Extraction et nettoyage de la Marque
        brand = "Inconnu"
        brand_el = soup.select_one("#bylineInfo")
        if brand_el:
            brand_raw = brand_el.text
            # Nettoyage des variantes textuelles d'Amazon
            brand_raw = re.sub(r"(Visitez la boutique|Marque\s*:|Brand\s*:)", "", brand_raw, flags=re.IGNORECASE)
            brand = brand_raw.strip()
        
        # 2. Extraction et nettoyage de la Note (juste le chiffre)
        rating = "Pas de note"
        rating_el = soup.select_one("span.a-icon-alt")
        if rating_el:
            rating_match = re.search(r"([0-9][,.]?[0-9]?)", rating_el.text)
            if rating_match:
                rating = rating_match.group(1).replace(",", ".")
        
        # 3. Extraction et nettoyage du Nombre d'avis (0 si vide, pas de parenthèses)
        reviews = "0"
        reviews_el = soup.select_one("#acrCustomerReviewText")
        if reviews_el:
            reviews_raw = reviews_el.text
            reviews_clean = re.sub(r"[() \s,.]", "", reviews_raw) # Enlève parenthèses, espaces, virgules
            reviews_match = re.search(r"(\d+)", reviews_clean)
            if reviews_match:
                reviews = reviews_match.group(1)
        
        # 4. Séparation des Bullet Points (1 par colonne, max 5)
        bullets_list = [li.text.strip() for li in soup.select("#feature-bullets ul li span.a-list-item")]
        # On s'assure d'avoir au moins 5 éléments dans la liste, même vides
        while len(bullets_list) < 5:
            bullets_list.append("")
        
        # 5. Nombre d'images
        images = set([img['src'] for img in soup.select("#altImages img") if "src" in img.attrs and ("overlay" not in img['src'])])
        nb_images = len(images) if images else 1
        
        # 6. Description
        desc_el = soup.select_one("#productDescription")
        desc = desc_el.text.strip() if desc_el else "Aucune description"
        
        # 7. Caractéristiques produits (Tableau technique)
        specs = {}
        for row in soup.select("#prodDetails table tr, #detailBullets_feature_div li"):
            th = row.select_one("th, span.a-text-bold")
            td = row.select_one("td, span:not(.a-text-bold)")
            if th and td:
                k = th.text.replace(":", "").strip()
                v = td.text.strip()
                if k and v:
                    specs[k] = v
        specs_str = str(specs) if specs else "Aucune caractéristique"
        
        # 8. RECHERCHE ULTRA-ROBUSTE DU BSR
        bsr_text = "Non trouvé"
        # Méthode A : Dans le texte global de la page (le plus fréquent)
        page_text = soup.get_text()
        bsr_match = re.search(r"Classement des meilleures ventes d'Amazon[\s\S]*?#([0-9\s,.]+) en", page_text, re.IGNORECASE)
        if bsr_match:
            bsr_text = f"#{bsr_match.group(1).strip()}"
        else:
            # Méthode B : Dans le dictionnaire des caractéristiques qu'on vient de scanner
            for key, val in specs.items():
                if "Classement des meilleures ventes" in key or "Classement" in key:
                    bsr_text = val
                    break

        # 9. Déclinaisons (Variantes)
        variants = [el.text.strip() for el in soup.select("#twister ul li span.a-button-text")]
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
        print(f"Erreur lors du scan de {asin} : {e}")
        return None

def executer_phase_2():
    chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)
    if not os.path.exists(chemin_fichier):
        print("Erreur : Le fichier historique_bestsellers.xlsx n'existe pas encore.")
        return

    df_fiches = pd.read_excel(chemin_fichier, sheet_name="Fiches_Produits")
    df_clst = pd.read_excel(chemin_fichier, sheet_name="Suivi_Classement")
    
    # On supprime la colonne Titre_Complet si elle traîne encore dans l'historique
    if "Titre_Complet" in df_fiches.columns:
        df_fiches = df_fiches.drop(columns=["Titre_Complet"])

    # Identification des lignes à scanner
    if "Note" not in df_fiches.columns:
        asins_a_scanner = df_fiches["ASIN"].tolist()
    else:
        asins_a_scanner = df_fiches[df_fiches["Marque"] == "À collecter en Phase 2"]["ASIN"].tolist()

    if not asins_a_scanner:
        print("Tous les produits sont déjà nettoyés et analysés ! Ras.")
        return

    # Lot de 15 pour rester sous le radar d'Amazon
    asins_a_scanner = asins_a_scanner[:15]
    print(f"Lancement du nettoyage profond sur {len(asins_a_scanner)} produits...")

    driver = configurer_driver()
    resultats = []
    
    try:
        for asin in asins_a_scanner:
            info = scraper_fiche_produit(driver, asin)
            if info:
                resultats.append(info)
            time.sleep(3)
    finally:
        driver.quit()

    if resultats:
        # Application des modifications dans le tableau Excel
        for res in resultats:
            asin = res["ASIN"]
            for col in res.keys():
                df_fiches.loc[df_fiches["ASIN"] == asin, col] = res[col]

        # Ré-écriture propre dans les deux onglets Excel
        with pd.ExcelWriter(chemin_fichier, engine="openpyxl") as writer:
            df_clst.to_excel(writer, sheet_name="Suivi_Classement", index=False)
            df_fiches.to_excel(writer, sheet_name="Fiches_Produits", index=False)
            
        print("Mise à jour et nettoyage de l'onglet 'Fiches_Produits' terminés avec succès !")
    else:
        print("Aucune donnée collectée.")

if __name__ == "__main__":
    executer_phase_2()
