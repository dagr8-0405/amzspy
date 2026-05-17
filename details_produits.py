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
    print(f"-> Scan du produit {asin}...")
    
    try:
        driver.get(url)
        time.sleep(4) # Pause de sécurité pour laisser charger les avis et caractéristiques
        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # 1. Titre Complet
        title_el = soup.select_one("#productTitle")
        title = title_el.text.strip() if title_el else "Inconnu"
        
        # 2. Marque
        brand_el = soup.select_one("#bylineInfo")
        brand = brand_el.text.replace("Visitez la boutique", "").replace("Marque :", "").strip() if brand_el else "Inconnu"
        
        # 3. Note et Avis
        rating_el = soup.select_one("span.a-icon-alt")
        rating = rating_el.text.split("sur")[0].strip() if rating_el else "Pas de note"
        
        reviews_el = soup.select_one("#acrCustomerReviewText")
        reviews = reviews_el.text.replace("évaluations", "").replace("évaluation", "").strip() if reviews_el else "0"
        
        # 4. Bullet Points (Tous dispos)
        bullets = [li.text.strip() for li in soup.select("#feature-bullets ul li span.a-list-item")]
        bullets_str = " | ".join(bullets) if bullets else "Aucun"
        
        # 5. Nombre d'images
        images = set([img['src'] for img in soup.select("#altImages img") if "src" in img.attrs and ("overlay" not in img['src'])])
        nb_images = len(images) if images else 1
        
        # 6. Description
        desc_el = soup.select_one("#productDescription")
        desc = desc_el.text.strip() if desc_el else "Aucune description textuelle"
        
        # 7. Caractéristiques produits (Tableau technique)
        specs = {}
        for row in soup.select("#prodDetails table tr"):
            th = row.select_one("th")
            td = row.select_one("td")
            if th and td:
                specs[th.text.strip()] = td.text.strip()
        specs_str = str(specs) if specs else "Aucune caractéristique trouvée"
        
        # 8. Classements & BSR additionnels
        bsr_text = "Inconnu"
        detail_text = soup.get_text()
        bsr_match = re.search(r"Classement des meilleures ventes d'Amazon([^\n]+)", detail_text)
        if bsr_match:
            bsr_text = bsr_match.group(0).strip()

        # 9. Déclinaisons (Variantes)
        variants = [el.text.strip() for el in soup.select("#twister ul li span.a-button-text")]
        nb_variants = len(variants)
        variants_str = ", ".join(variants) if variants else "Aucune déclinaison"

        return {
            "ASIN": asin,
            "Titre_Complet": title,
            "Marque": brand,
            "Note": rating,
            "Nb_Avis": reviews,
            "Bullet_Points": bullets_str,
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
        print("Erreur : Le fichier historique_bestsellers.xlsx n'existe pas encore. Lance d'abord la Phase 1.")
        return

    # On charge l'onglet des fiches produits
    df_fiches = pd.read_excel(chemin_fichier, sheet_name="Fiches_Produits")
    df_clst = pd.read_excel(chemin_fichier, sheet_name="Suivi_Classement")
    
    # On cherche les ASINs qui ont besoin d'être scannés (ceux qui contiennent encore "À collecter en Phase 2")
    # Ou ceux dont les colonnes indispensables sont absentes
    if "Note" not in df_fiches.columns:
        asins_a_scanner = df_fiches["ASIN"].tolist()
    else:
        asins_a_scanner = df_fiches[df_fiches["Marque"] == "À collecter en Phase 2"]["ASIN"].tolist()

    if not asins_a_scanner:
        print("Tous les produits sont déjà analysés à 100 % ! Rien à faire.")
        return

    print(f"Trouvé : {len(asins_a_scanner)} nouveaux produits uniques à analyser.")
    
    # On limite à 15 produits par session pour ne pas saturer GitHub (le reste se fera aux prochains lancements)
    asins_a_scanner = asins_a_scanner[:15]
    print(f"Lancement du scan pour un lot de {len(asins_a_scanner)} produits...")

    driver = configurer_driver()
    resultats = []
    
    try:
        for asin in asins_a_scanner:
            info = scraper_fiche_produit(driver, asin)
            if info:
                resultats.append(info)
            time.sleep(3) # Sécurité anti-blocage
    finally:
        driver.quit()

    if resultats:
        df_nouveaux_details = pd.DataFrame(resultats)
        
        # Fusionner les anciennes données avec les nouvelles informations récoltées
        for res in resultats:
            asin = res["ASIN"]
            # Si l'ASIN existe, on remplace la ligne par les vraies infos
            for col in res.keys():
                df_fiches.loc[df_fiches["ASIN"] == asin, col] = res[col]

        # Sauvegarde des deux onglets
        with pd.ExcelWriter(chemin_fichier, engine="openpyxl") as writer:
            df_clst.to_excel(writer, sheet_name="Suivi_Classement", index=False)
            df_fiches.to_excel(writer, sheet_name="Fiches_Produits", index=False)
            
        print("Mise à jour de l'onglet 'Fiches_Produits' réussie !")
    else:
        print("Aucune donnée n'a pu être collectée lors de cette session.")

if __name__ == "__main__":
    executer_phase_2()
