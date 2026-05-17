import os
import re
import time
from datetime import datetime
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

URL = "https://www.amazon.fr/gp/bestsellers/electronics/15855785031/"
NOM_FICHIER_CENTRAL = "historique_bestsellers.xlsx"

def scraper_amazon_onglets():
    print(f"[{datetime.now().strftime('%H:%M:%S')}] Phase 1 : Extraction du classement...")

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    # Force la langue française au niveau du navigateur
    chrome_options.add_argument("--lang=fr-FR")
    chrome_options.add_experimental_option('prefs', {'intl.accept_languages': 'fr,fr-FR'})
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        driver.get(URL)
        time.sleep(6)
        for i in range(1, 6):
            driver.execute_script(f"window.scrollTo(0, (document.body.scrollHeight / 5) * {i});")
            time.sleep(2)
        soup = BeautifulSoup(driver.page_source, "html.parser")
    finally:
        driver.quit()

    liste_classement = []
    liste_produits_uniques = []
    elements = soup.select("div[id^='p13n-asin-index-'], div.p13n-grid-content, div.zg-grid-general-faceout, li.zg-item-immersion")

    asins_vus_ce_jour = set()

    for index, el in enumerate(elements, 1):
        try:
            rank_el = el.select_one("span.zg-badge-text, span.p13n-sc-zg-badge-text, span.scr-zg-badge-text")
            rank = rank_el.text.strip() if rank_el else f"#{index}"

            title = "Titre inconnu"
            title_el = el.select_one("div._cDE42_title_3D69b, div.p13n-sc-truncate-desktop-type2, div.p13n-sc-css-line-clamp-2")
            if title_el:
                title = title_el.text.strip()

            price_pure = 0.0
            price_el = el.select_one("span.a-offscreen, span.p13n-sc-price")
            if price_el:
                price_text = price_el.text.replace("€", "").replace(",", ".").strip()
                price_match = re.search(r"([0-9.]+)", price_text)
                if price_match:
                    price_pure = float(price_match.group(1))

            asin = "Inconnu"
            link_el = el.select_one("a[href*='/dp/'], a[href*='/gp/product/']")
            if link_el and "href" in link_el.attrs:
                asin_match = re.search(r"/(?:dp|product)/([A-Z0-9]{10})", link_el["href"])
                if asin_match:
                    asin = asin_match.group(1)

            # ANTI-DOUBLON : On vérifie si l'ASIN n'a pas déjà été attrapé ce matin
            if asin != "Inconnu" and asin not in asins_vus_ce_jour:
                asins_vus_ce_jour.add(asin)
                
                liste_classement.append({
                    "Date": datetime.now().strftime("%Y-%m-%d"),
                    "Classement": rank,
                    "ASIN": asin,
                    "Prix": price_pure
                })
                
                liste_produits_uniques.append({
                    "ASIN": asin,
                    "Titre": title,
                    "Marque": "À collecter en Phase 2",
                    "Date_Detection": datetime.now().strftime("%Y-%m-%d")
                })
        except Exception:
            continue

    if liste_classement:
        df_nouveau_clst = pd.DataFrame(liste_classement)
        df_nouvelles_fiches = pd.DataFrame(liste_produits_uniques)

        chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)

        if os.path.exists(chemin_fichier):
            try:
                df_ancien_clst = pd.read_excel(chemin_fichier, sheet_name="Suivi_Classement")
                df_final_clst = pd.concat([df_ancien_clst, df_nouveau_clst], ignore_index=True).drop_duplicates(subset=["Date", "Classement"])
            except Exception:
                df_final_clst = df_nouveau_clst

            try:
                df_anciennes_fiches = pd.read_excel(chemin_fichier, sheet_name="Fiches_Produits")
                df_final_fiches = pd.concat([df_anciennes_fiches, df_nouvelles_fiches], ignore_index=True).drop_duplicates(subset=["ASIN"], keep="first")
            except Exception:
                df_final_fiches = df_nouvelles_fiches
        else:
            df_final_clst = df_nouveau_clst
            df_final_fiches = df_nouvelles_fiches

        with pd.ExcelWriter(chemin_fichier, engine="openpyxl") as writer:
            df_final_clst.to_excel(writer, sheet_name="Suivi_Classement", index=False)
            df_final_fiches.to_excel(writer, sheet_name="Fiches_Produits", index=False)
        print("Phase 1 Terminée avec succès.")
