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
    print(
        f"[{datetime.now().strftime('%H:%M:%S')}] Lancement du navigateur sécurisé (Mode Double Onglet)..."
    )

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Connexion à Amazon...")
        driver.get(URL)
        time.sleep(5)

        for i in range(1, 6):
            driver.execute_script(
                f"window.scrollTo(0, (document.body.scrollHeight / 5) * {i});"
            )
            time.sleep(2.5)

        soup = BeautifulSoup(driver.page_source, "html.parser")

    finally:
        driver.quit()

    liste_classement = []
    liste_produits_uniques = []
    elements = soup.select(
        "div[id^='p13n-asin-index-'], div.p13n-grid-content, div.zg-grid-general-faceout, li.zg-item-immersion"
    )

    print(f"Analyse de {len(elements)} emplacements...")

    for index, el in enumerate(elements, 1):
        try:
            # 1. Extraction du Rang
            rank_el = el.select_one(
                "span.zg-badge-text, span.p13n-sc-zg-badge-text, span.scr-zg-badge-text"
            )
            rank = rank_el.text.strip() if rank_el else f"#{index}"

            # 2. Extraction du Titre
            title = "Titre inconnu"
            title_el = el.select_one(
                "div._cDE42_title_3D69b, div.p13n-sc-truncate-desktop-type2, div.p13n-sc-css-line-clamp-2, div.p13n-sc-css-line-clamp-1"
            )
            if title_el:
                title = title_el.text.strip()

            if title == "Titre inconnu" or title == "Regarder" or "€" in title:
                for target in el.select("span, div, a"):
                    txt = target.text.strip()
                    if (
                        len(txt) > 25
                        and "€" not in txt
                        and txt != "Regarder"
                        and "étoiles" not in txt
                    ):
                        title = txt
                        break

            # 3. Extraction du Prix
            price = "Non communiqué"
            price_el = el.select_one("span.a-offscreen")
            if price_el:
                price = price_el.text.strip()
            else:
                alt_price = el.select_one(
                    "span.p13n-sc-price, span.a-color-price"
                )
                if alt_price:
                    price = alt_price.text.strip()

            # 4. Extraction de l'ASIN
            asin = "Inconnu"
            link_el = el.select_one("a[href*='/dp/'], a[href*='/gp/product/']")
            if link_el and "href" in link_el.attrs:
                href = link_el["href"]
                asin_match = re.search(r"/dp/([A-Z0-9]{10})", href)
                if not asin_match:
                    asin_match = re.search(r"/gp/product/([A-Z0-9]{10})", href)
                if asin_match:
                    asin = asin_match.group(1)

            if title and title != "Titre inconnu" and "€" not in title and asin != "Inconnu":
                # Données pour l'onglet Classement
                liste_classement.append(
                    {
                        "Date": datetime.now().strftime("%Y-%m-%d"),
                        "Classement": rank,
                        "ASIN": asin,
                        "Prix": price,
                    }
                )
                # Données temporaires pour l'onglet Fiches Produits
                liste_produits_uniques.append(
                    {
                        "ASIN": asin,
                        "Titre": title,
                        "Marque": "À collecter en Phase 2",
                        "Description": "À collecter en Phase 2",
                        "Attributs": "À collecter en Phase 2",
                        "Date_Detection": datetime.now().strftime("%Y-%m-%d"),
                    }
                )
        except Exception:
            continue

    if liste_classement:
        # Nettoyage du classement du jour
        df_nouveau_clst = pd.DataFrame(liste_classement)
        df_nouveau_clst["numeric_rank"] = (
            df_nouveau_clst["Classement"].str.extract(r"(\d+)").astype(float)
        )
        df_nouveau_clst = df_nouveau_clst.drop_duplicates(
            subset=["numeric_rank"]
        )
        df_nouveau_clst = df_nouveau_clst.sort_values(by="numeric_rank")
        df_nouveau_clst["Classement"] = df_nouveau_clst["numeric_rank"].apply(
            lambda x: f"#{int(x)}" if pd.notnull(x) else ""
        )
        df_nouveau_clst = df_nouveau_clst.drop(columns=["numeric_rank"])

        df_nouvelles_fiches = pd.DataFrame(liste_produits_uniques).drop_duplicates(subset=["ASIN"])

        chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)

        # Chargement de l'existant ou création
        if os.path.exists(chemin_fichier):
            try:
                df_ancien_clst = pd.read_excel(
                    chemin_fichier, sheet_name="Suivi_Classement"
                )
                df_final_clst = pd.concat(
                    [df_ancien_clst, df_nouveau_clst], ignore_index=True
                )
            except Exception:
                df_final_clst = df_nouveau_clst

            try:
                df_anciennes_fiches = pd.read_excel(
                    chemin_fichier, sheet_name="Fiches_Produits"
                )
                # Fusion intelligente : On ne garde que les ASINs qui n'existaient pas déjà !
                df_final_fiches = pd.concat(
                    [df_anciennes_fiches, df_nouvelles_fiches],
                    ignore_index=True,
                ).drop_duplicates(subset=["ASIN"], keep="first")
            except Exception:
                df_final_fiches = df_nouvelles_fiches
        else:
            df_final_clst = df_nouveau_clst
            df_final_fiches = df_nouvelles_fiches

        # Écriture dans les deux onglets Excel
        with pd.ExcelWriter(chemin_fichier, engine="openpyxl") as writer:
            df_final_clst.to_excel(
                writer, sheet_name="Suivi_Classement", index=False
            )
            df_final_fiches.to_excel(
                writer, sheet_name="Fiches_Produits", index=False
            )

        print(
            f"Sauvegarde réussie dans '{NOM_FICHIER_CENTRAL}' (Structure Double Onglet OK)."
        )
    else:
        print("Erreur : Aucun produit trouvé.")


if __name__ == "__main__":
    scraper_amazon_onglets()
