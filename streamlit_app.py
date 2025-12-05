import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Comparaison d'inventaires par agence", layout="wide")

st.title("Comparaison d'écarts d'inventaire entre agences")

st.markdown(
    """
Cette application permet de :
- Charger les fichiers d'inventaire de plusieurs agences (au moins 2),
- Saisir un **nom ou code d'agence** pour chaque fichier,
- Comparer les **écarts déjà calculés** (colonne *Ecart* dans les fichiers),
- Générer **un fichier Excel par agence**, avec **un onglet par agence comparée**.
"""
)

# --- Gestion dynamique du nombre d'agences ---

if "nb_agences" not in st.session_state:
    st.session_state.nb_agences = 2  # minimum 2 agences

col_left, col_right = st.columns([1, 3])
with col_left:
    if st.button("+ Ajouter une agence"):
        st.session_state.nb_agences += 1

st.write(f"Nombre d'agences à comparer : **{st.session_state.nb_agences}**")

# --- Fonction utilitaire pour charger et normaliser un inventaire ---

def charger_inventaire(uploaded_file: BytesIO) -> pd.DataFrame:
    """
    Lit un fichier d'inventaire (CSV ; ou Excel) et renvoie un DataFrame
    normalisé avec les colonnes :
    - Code article
    - Désignation
    - Stock théorique
    - Stock physique
    - Ecart (déjà calculé dans le fichier)
    """
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file, sep=";", encoding="latin1")
    else:
        df = pd.read_excel(uploaded_file)

    # Nettoyage des noms de colonnes (espaces, NBSP, etc.)
    df.columns = [c.replace("\xa0", " ").strip() for c in df.columns]

    # On suppose une structure identique au fichier AIN
    required_logical_cols = [
        "Code article",
        "Désignation",
        "Stock théorique",
        "Stock physique",
        "Ecart",
    ]

    # On détecte les colonnes même si les libellés varient légèrement
    col_map = {}
    for col in df.columns:
        col_clean = col.lower().replace("\xa0", " ")
        if "code article" in col_clean and "Code article" not in col_map:
            col_map["Code article"] = col
        elif "désignation" in col_clean and "Désignation" not in col_map:
            col_map["Désignation"] = col
        elif "stock théorique" in col_clean and "Stock théorique" not in col_map:
            col_map["Stock théorique"] = col
        elif "stock physique" in col_clean and "Stock physique" not in col_map:
            col_map["Stock physique"] = col
        elif col_clean.startswith("ecart") and "Ecart" not in col_map:
            col_map["Ecart"] = col

    missing = [c for c in required_logical_cols if c not in col_map]
    if missing:
        raise ValueError(
            f"Colonnes manquantes ou non reconnues dans le fichier : {missing}\n"
            f"Colonnes trouvées : {list(df.columns)}"
        )

    df_inv = df[
        [
            col_map["Code article"],
            col_map["Désignation"],
            col_map["Stock théorique"],
            col_map["Stock physique"],
            col_map["Ecart"],
        ]
    ].copy()

    df_inv.columns = [
        "Code article",
        "Désignation",
        "Stock théorique",
        "Stock physique",
        "Ecart",   # on garde l'écart tel quel (déjà calculé)
    ]

    # Conversion numérique (on ne recalcule pas l'écart, on le convertit juste en nombre)
    for col in ["Stock théorique", "Stock physique", "Ecart"]:
        df_inv[col] = (
            df_inv[col]
            .astype(str)
            .str.replace("\xa0", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        df_inv[col] = pd.to_numeric(df_inv[col], errors="coerce")

    # On supprime les lignes sans code article
    df_inv = df_inv.dropna(subset=["Code article"])
    df_inv["Code article"] = df_inv["Code article"].astype(str).str.strip()

    return df_inv


def comparer_deux_agences(nom_a, df_a, nom_b, df_b) -> pd.DataFrame:
    """
    Compare les écarts d'inventaire entre deux agences, sur les références communes.
    Utilise la colonne Ecart déjà présente dans chaque fichier.

    Retourne un DataFrame avec :
    - Agence (référence)
    - Agence comparée
    - Code article
    - Désignation (de l'agence A)
    - Stock théorique (de A)
    - Stock physique (de A)
    - Ecart_<A> (issu du fichier A)
    - Ecart_<B> (issu du fichier B)
    - Ecart_des_2_agences = Ecart_B - Ecart_A
    - Somme_ecarts = Ecart_A + Ecart_B
    """
    # On renomme la colonne Ecart pour chaque agence
    a = df_a.copy()
    b = df_b.copy()
    a = a.rename(columns={"Ecart": f"Ecart_{nom_a}"})
    b = b.rename(columns={"Ecart": f"Ecart_{nom_b}"})

    # Jointure sur les codes article (références communes)
    merged = a.merge(
        b[["Code article", f"Ecart_{nom_b}"]],
        on="Code article",
        how="inner",
    )

    res = pd.DataFrame(
        {
            "Agence": nom_a,
            "Agence comparée": nom_b,
            "Code article": merged["Code article"],
            "Désignation": merged["Désignation"],
            "Stock théorique": merged["Stock théorique"],
            "Stock physique": merged["Stock physique"],
            f"Ecart_{nom_a}": merged[f"Ecart_{nom_a}"],
            f"Ecart_{nom_b}": merged[f"Ecart_{nom_b}"],
        }
    )

    # Calculs d'écarts à partir des écarts fournis dans les fichiers
    res["Ecart_des_2_agences"] = res[f"Ecart_{nom_b}"] - res[f"Ecart_{nom_a}"]
    res["Somme_ecarts"] = res[f"Ecart_{nom_a}"] + res[f"Ecart_{nom_b}"]

    # Tri par valeur absolue de Somme_ecarts décroissante (optionnel)
    res = res.sort_values(by="Somme_ecarts", key=lambda s: s.abs(), ascending=False)

    return res


# --- Saisie des fichiers & noms d'agences ---

st.subheader("Chargement des inventaires agences")

agence_infos = []  # liste de tuples (nom_agence, df_inventaire)

for i in range(st.session_state.nb_agences):
    st.markdown(f"### Agence {i + 1}")
    col1, col2 = st.columns([2, 3])

    with col1:
        nom_agence = st.text_input(
            f"Nom ou code de l'agence {i + 1}",
            key=f"nom_agence_{i}",
        )

    with col2:
        uploaded = st.file_uploader(
            f"Fichier d'inventaire (CSV ; ou Excel) pour l'agence {i + 1}",
            type=["csv", "xls", "xlsx"],
            key=f"file_agence_{i}",
        )

    if uploaded is not None and nom_agence.strip():
        try:
            df_inv = charger_inventaire(uploaded)
            agence_infos.append((nom_agence.strip(), df_inv))

            with st.expander(f"Aperçu des premières lignes - {nom_agence}", expanded=False):
                st.dataframe(df_inv.head(10))
        except Exception as e:
            st.error(f"Erreur lors du chargement de l'agence {i + 1} : {e}")

st.markdown("---")

if st.button("Générer les rapports Excel"):
    if len(agence_infos) < 2:
        st.error("Il faut au moins **2 agences** valides (nom + fichier) pour lancer la comparaison.")
    else:
        st.success("Comparaison en cours…")

        # Pour chaque agence, on génère un fichier Excel avec un onglet par agence comparée
        for idx_ref, (nom_ref, df_ref) in enumerate(agence_infos):
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                for idx_other, (nom_other, df_other) in enumerate(agence_infos):
                    if idx_other == idx_ref:
                        continue  # on ne se compare pas à soi-même

                    df_comp = comparer_deux_agences(nom_ref, df_ref, nom_other, df_other)

                    # Nom d'onglet : <AgenceRef>_vs_<AgenceAutre> (tronqué à 31 caractères)
                    sheet_name = f"{nom_ref}_vs_{nom_other}"
                    sheet_name = sheet_name[:31]  # limite Excel

                    df_comp.to_excel(writer, index=False, sheet_name=sheet_name)

            buffer.seek(0)

            st.download_button(
                label=f"📥 Télécharger le rapport pour l'agence {nom_ref}",
                data=buffer,
                file_name=f"rapport_inventaire_{nom_ref}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.info(
            "Un fichier Excel a été généré **pour chaque agence**, "
            "avec un onglet par agence comparée contenant les références communes et les écarts."
        )
