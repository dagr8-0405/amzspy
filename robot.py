from datetime import datetime
import os
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# URL cible (Boîtiers Montres Amazon)
URL = "https://www.amazon.fr/gp/bestsellers/electronics/15855785031/"


def scraper_amazon_cloud():
    print(
        f"[{datetime.now().strftime('%H:%M:%S')}] Lancement du navigateur sécurisé (Mode Cloud)..."
    )

    chrome_options = Options()
    # Configuration indispensable pour tourner sur les serveurs de GitHub sans écran
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
        time.sleep(6)

        # Descente très progressive en 5 étapes pour forcer l'affichage des 100 produits
        for i in range(1, 6):
            driver.execute_script(
                f"window.scrollTo(0, (document.body.scrollHeight / 5) * {i});"
            )
            time.sleep(2.5)

        soup = BeautifulSoup(driver.page_source, "html.parser")

    finally:
        driver.quit()

    liste_produits = []

    # Sélection ultra-large de toutes les boîtes de la grille Amazon (gauche, droite, paires, impaires)
    elements = soup.select(
        "div[id^='p13n-asin-index-'], div.p13n-grid-content, div.zg-grid-general-faceout, li.zg-item-immersion"
    )

    print(
        f"Analyse en cours de {len(elements)} emplacements trouvés sur la page..."
    )

    for index, el in enumerate(elements, 1):
        try:
            # 1. Extraction du Rang (Classement)
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

            # Sécurité anti-bug pour les titres
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

            # 3. Extraction du Prix (Sélecteur universel)
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

            # On valide la ligne si le titre est cohérent
            if title and title != "Titre inconnu" and "€" not in title:
                liste_produits.append(
                    {
                        "Date": datetime.now().strftime("%Y-%m-%d"),
                        "Classement": rank,
                        "Titre": title,
                        "Prix": price,
                    }
                )
        except Exception:
            continue

    # Traitement et mise en forme du tableau final
    if liste_produits:
        df = pd.DataFrame(liste_produits)

        # Nettoyage et extraction du numéro du classement
        df["numeric_rank"] = (
            df["Classement"].str.extract(r"(\d+)").astype(float)
        )

        # Suppression des doublons de place
        df = df.drop_duplicates(subset=["numeric_rank"])

        # Tri mathématique parfait (1, 2, 3, 4...)
        df = df.sort_values(by="numeric_rank")

        # Reconstruction propre de la colonne Classement (ex: #1, #2, #3...)
        df["Classement"] = df["numeric_rank"].apply(
            lambda x: f"#{int(x)}" if pd.notnull(x) else ""
        )
        df = df.drop(columns=["numeric_rank"])

        # Création du fichier Excel avec la date du jour
        date_str = datetime.now().strftime("%Y-%m-%d")
        nom_fichier = f"amazon_bestsellers_{date_str}.xlsx"

        df.to_excel(os.path.join(os.getcwd(), nom_fichier), index=False)
        print(
            f"Succès total ! Le fichier '{nom_fichier}' a été créé dans le cloud sans aucune perte."
        )
    else:
        print("Erreur : Aucun produit trouvé.")


if __name__ == "__main__":
    scraper_amazon_cloud()
