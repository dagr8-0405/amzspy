import os
import re
import pandas as pd

NOM_FICHIER_CENTRAL = "historique_bestsellers.xlsx"
SCRIPT_TARGET = "details_produits.py"

def verifier_coherence_donnees():
    print("🔍 [IA] Analyse de contrôle de la lisibilité des données...")
    chemin_fichier = os.path.join(os.getcwd(), NOM_FICHIER_CENTRAL)
    
    if not os.path.exists(chemin_fichier):
        print("❌ Fichier Excel introuvable. Impossible de vérifier.")
        return

    # Chargement de l'onglet des fiches
    df_fiches = pd.read_excel(chemin_fichier, sheet_name="Fiches_Produits")
    
    anomalies = []
    
    # On parcourt les lignes récemment scannées (celles qui ont une Date_Analyse)
    df_recent = df_fiches[df_fiches["Date_Analyse"].notna()]
    
    for _, row in df_recent.iterrows():
        asin = row["ASIN"]
        raison = []
        
        # RÈGLE 1 : La marque ne doit pas être Inconnu
        if str(row["Marque"]).lower() in ["inconnu", "nan", ""]:
            raison.append("Marque non récupérée (Inconnu)")
            
        # RÈGLE 2 : La note doit être un chiffre pur (pas 0.0 sauf si vrai produit sans avis)
        if row["Note"] == 0.0 and row["Nb_Avis"] > 0:
            raison.append("Note bloquée à 0.0 alors que le produit a des avis")
            
        # RÈGLE 3 : Le BSR ne doit pas être 'Non trouvé'
        if str(row["BSR_Categories"]).lower() in ["non trouvé", "nan", ""]:
            raison.append("BSR manquant (Non trouvé)")

        if raison:
            anomalies.append({
                "ASIN": asin,
                "Erreurs": ", ".join(raison),
                "Caracteristiques_Brutes": row.get("Caracteristiques", "Aucune")
            })

    if anomalies:
        print(f"⚠️ [IA] {len(anomalies)} produit(s) présente(nt) des défauts de lecture.")
        creer_ticket_anomalie(anomalies[0]) # On traite le premier bloqueur pour corriger le code
    else:
        print("✅ [IA] Félicitations ! 100% des données récoltées sont lisibles et conformes. Pas d'auto-correction requise.")

def creer_ticket_anomalie(anomalie):
    # Ce fichier texte sera lu par l'Action GitHub pour ouvrir une Issue automatique
    with open("alerte_ia.txt", "w", encoding="utf-8") as f:
        f.write(f"ASIN CIBLE : {anomalie['ASIN']}\n")
        f.write(f"ANOMALIE DETECTEE : {anomalie['Erreurs']}\n")
        f.write(f"DONNEES BRUTES RECUES : {anomalie['Caracteristiques_Brutes']}\n")
    print("📝 [IA] Fichier d'alerte 'alerte_ia.txt' généré pour l'auto-correction.")

if __name__ == "__main__":
    verifier_coherence_donnees()
