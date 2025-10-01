# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook


st.title(" Transformation des commandes pizza")

# Étape 1 : Upload du fichier Excel
fichier = st.file_uploader("Importe le fichier Excel brut (.xlsx)", type=["xlsx"])

if fichier is not None:
    # Charger le fichier dans un DataFrame
    df = pd.read_excel(fichier)

    # ⚡ Exemple de traitement (à remplacer par ta vraie fonction)
    def heure(row):
        if "17-18" in row["SKU"]:
            return "18h - 18h30"
        if "18-19" in row["SKU"]:
            return "18h30 - 19h30"
        if "19-20" in row["SKU"]:
            return "19h30 - 20h30"
        if "20-21" in row["SKU"]:
            return "20h30 - 21h30"

    df['No de commande'] = df['Order Number']
    df['Nom'] = df['First Name (Billing)']
    df['Prénom'] = df["Last Name (Billing)"]
    df['Adresse'] = df['Address 1&2 (Billing)']
    df['Téléphone'] = df['Phone (Billing)']
    df['Heure de livraison'] = df.apply(heure, axis=1)
    df["Nom item"]=df["Item Name"]
    df["remarque"]=df["Customer Note"]

    def hawai (row):
        if "hawa" in row["Nom item"].lower() :
            return row["Quantity"]
    def prosciuto (row):
        if ("pr") in row["Nom item"].lower() :
            if not "fu" in row["Nom item"].lower():
                return row["Quantity"]
    def prosciuto_ef (row):
        if ("pr" and "fu") in row["Nom item"].lower() :
            return row["Quantity"]

    def margherita (row):
        if "marg" in row["Nom item"].lower() :
            return row["Quantity"]

    

    df['hawai'] = df.apply(hawai, axis=1).fillna(0)
    df["jambon"] = df.apply(prosciuto, axis=1).astype(float).fillna(0)
    df["margherita"] = df.apply(margherita, axis =1).astype(float).fillna(0)
    df["jambon_champi"] = df.apply(prosciuto_ef, axis=1).fillna(0)
    df["prix_hawai_14"] = df["hawai"] * 14
    df["prix_jambon_14"] = df["jambon"] * 14
    df["prix_margherita_12"] = df["margherita"] * 12
    df["prix_jambon_champi_14"] = df["jambon_champi"] *12
    df["prix total"] = df["prix_hawai_14"] + df["prix_jambon_14"] + df["prix_margherita_12"] + df["prix_jambon_champi_14"]



    df_new = (
        df.groupby("No de commande")
        .agg({
            "Nom": "first",
            "Prénom": "first",
            "Adresse": "first", 
            "Téléphone": "first",            # même valeur partout → on garde la 1ère
            "Heure de livraison": "first",       # idem
            "hawai": "sum",
            "prix_hawai_14" : "sum",
            "jambon": "sum",
            "prix_jambon_14" : "sum",        # additionne les quantités
            "margherita": "sum",
            "prix_margherita_12" : "sum",
            "jambon_champi": "sum",
            "prix_jambon_champi_14": "sum", 
            "prix total": "sum",
            "remarque": "sum",
        })
        .reset_index()
    )
    zerop=["remarque"]
    df_new[zerop] = df_new[zerop].replace(0, "")
    df_new.to_excel("v1.xlsx", index=False)

    from openpyxl import load_workbook
    wb = load_workbook("v1.xlsx")
    s = wb.active



# Adapter la largeur des colonnes
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # lettre de la colonne
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 2  # +2 pour un peu d'espace

     wb.save("commandes_finales.xlsx")
    # Étape 2 : Afficher un aperçu
    st.write("Aperçu du fichier transformé 👇")
    st.dataframe(df.head())

  

    # Étape 3 : Permettre le téléchargement
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)

    st.download_button(
        label="⬇️ Télécharger le fichier transformé",
        data=buffer,
        file_name="commandes_traitees.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
