import os
import pandas as pd

NOM_FICHIER_CENTRAL = "historique_bestsellers.xlsx"

def verifier_coherence_donnees():
    print("🔍 [IA] Analyse de contrôle de la lisibilité des données...")
    chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)
    
    if not os.path.exists(chemin_fichier):
        return

    # Si le fichier d'alerte existe déjà (généré par le captcha), on s'arrête là pour le traiter
    if os.path.exists("alerte_ia.txt"):
        print("⚠️ [IA] Un blocage de sécurité en amont a déjà été identifié.")
        return

    df_fiches = pd.read_excel(chemin_fichier, sheet_name="Fiches_Produits")
    df_recent = df_fiches[df_fiches["Date_Analyse"].notna()]
    
    anomalies = []
    for _, row in df_recent.iterrows():
        asin = row["ASIN"]
        raison = []
        
        if str(row["Marque"]).lower() in ["inconnu", "nan", ""]:
            raison.append("Marque manquante")
        if str(row["BSR_Categories"]).lower() in ["non trouvé", "nan", ""]:
            raison.append("BSR manquant")

        if raison:
            anomalies.append({
                "ASIN": asin,
                "Erreurs": ", ".join(raison),
                "Brut": row.get("Caracteristiques", "Aucune")
            })

    if anomalies:
        print(f"⚠️ [IA] {len(anomalies)} produit(s) non conforme(s).")
        with open("alerte_ia.txt", "w", encoding="utf-8") as f:
            f.write(f"ASIN CIBLE : {anomalies[0]['ASIN']}\n")
            f.write(f"ANOMALIE DETECTEE : {anomalies[0]['Erreurs']}\n")
            f.write(f"DONNEES BRUTES : {anomalies[0]['Brut']}\n")
    else:
        print("✅ [IA] Données conformes.")

if __name__ == "__main__":
    verifier_coherence_donnees()
