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


def scraper_amazon_centralise():
    print(
        f"[{datetime.now().strftime('%H:%M:%S')}] Lancement du navigateur sécurisé..."
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

    liste_produits = []
    elements = soup.select(
        "div[id^='p13n-asin-index-'], div.p13n-grid-content, div.zg-grid-general-faceout, li.zg-item-immersion"
    )

    print(
        f"Analyse en cours de {len(elements)} emplacements trouvés sur la page..."
    )

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

            # 4. EXTRACTION DE L'ASIN (L'identifiant produit indispensable pour la suite)
            asin = "Inconnu"
            # On cherche l'ASIN dans les liens du produit
            link_el = el.select_one("a[href*='/dp/'], a[href*='/gp/product/']")
            if link_el and "href" in link_el.attrs:
                href = link_el["href"]
                asin_match = re.search(r"/dp/([A-Z0-9]{10})", href)
                if not asin_match:
                    asin_match = re.search(r"/gp/product/([A-Z0-9]{10})", href)
                if asin_match:
                    asin = asin_match.group(1)

            if title and title != "Titre inconnu" and "€" not in title:
                liste_produits.append(
                    {
                        "Date": datetime.now().strftime("%Y-%m-%d"),
                        "Classement": rank,
                        "ASIN": asin,
                        "Titre": title,
                        "Prix": price,
                    }
                )
        except Exception:
            continue

    if liste_produits:
        df_nouveau = pd.DataFrame(liste_produits)
        df_nouveau["numeric_rank"] = (
            df_nouveau["Classement"].str.extract(r"(\d+)").astype(float)
        )
        df_nouveau = df_nouveau.drop_duplicates(subset=["numeric_rank"])
        df_nouveau = df_nouveau.sort_values(by="numeric_rank")
        df_nouveau["Classement"] = df_nouveau["numeric_rank"].apply(
            lambda x: f"#{int(x)}" if pd.notnull(x) else ""
        )
        df_nouveau = df_nouveau.drop(columns=["numeric_rank"])

        # GESTION DU FICHIER UNIQUE : On ajoute à l'existant sans écraser
        chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)

        if os.path.exists(chemin_fichier):
            # Si le fichier existe déjà, on charge l'ancien et on colle le nouveau en dessous
            df_ancien = pd.read_excel(chemin_fichier)
            df_final = pd.concat([df_ancien, df_nouveau], ignore_index=True)
            print(f"Fichier central complété. Total : {len(df_final)} lignes.")
        else:
            # Sinon, on crée le tout premier bloc
            df_final = df_nouveau
            print(f"Création du fichier central initial.")

        df_final.to_excel(chemin_fichier, index=False)
        print(f"Sauvegarde réussie dans '{NOM_FICHIER_CENTRAL}'.")
    else:
        print("Erreur : Aucun produit trouvé.")


if __name__ == "__main__":
    scraper_amazon_centralise()
