import sys
import streamlit as st
import pandas as pd
import pprint
from datetime import datetime
from pathlib import Path
import base64
from openpyxl import load_workbook
import numpy as np
import openpyxl
import pandas as pd
import re
from datetime import datetime, time, timedelta



def normalize_product_name(name):
    if not name:
        return ""
    # Mise en majuscules et suppression des caractères spéciaux
    name = name.upper()
    
    # Supprime les espaces entre les lettres : "D A P" → "DAP"
    name = re.sub(r'\b([A-Z])\s+([A-Z])\b', r'\1\2', name)

    # Supprime tous les caractères non alphanumériques et remplace-les par un espace
    name = re.sub(r'[^A-Z0-9]', ' ', name)

    # Supprime les espaces multiples
    name = re.sub(r'\s+', ' ', name).strip()

    # Dictionnaire des règles de normalisation
    patterns = {
        'MAP 11 52 SPECIAL': [
            r'MAP\s*11\s*52',
            r'MAP\s*SPECIAL',
            r'MAP\s*11\s*52\s*SPECIAL',
            r'MAP\s*11\s*52\s*Special\s*Low\s*Cd',

        ],
        'NPK 14 18 18 6S 1B2O3': [
            r'NPK\s*14\s*18\s*18\s*6S\s*1B2O3\s*AFRIQUE',
            r'NPK\s*14\s*18\s*18\s*6S\s*1B2O3'],
         

        'DAP SPECIAL': [
            r'DAP\s*SPC',
            r'DAP\s*SPECIAL',
        ],
        'DAP EURO': [
            r'DAP\s*EURO',
            r'DAP\s*EU',
            r'DAP\s*Euro\s*Low\s*Cd'
            
        ],
        'DAP STANDARD': [
            r'DAP\s*STANDARD',
            r'DAP\s*STD',
        ],
        'TSP SPECIAL JORF': [
            r'TSP\s*JORF',
            r'TSP-JORF',
            r'TSP\s*SPECIAL\s*JORF',
            r'TSP\s*SPC\s*JARF',

        ],
        'NPS 3 30 9S': [
            r'NPS\s*3\s*30\s*9S\s*OFAS',
            r'NPS\s*3\s*30\s*9S',
        ],
        'UREE': [r'\bUREE\b', r'\bURÉE\b']
    }

    for standard, variants in patterns.items():
        for pattern in variants:
            if re.search(pattern, name):
                return standard

    return name# retourne le nom nettoyé mais non reconnu


# Configurer la mise en page
st.set_page_config(
    layout="wide",
    page_title="Suivi des chargements",
    page_icon="🌐",
)

# ✅ CSS : Barre verte + placement heure à droite
st.markdown("""
    <style>
        .top-bar {
            background-color: #00a65a; /* Vert */
            height: 4px;
            width: 100%;
            position: fixed;
            top: 0;
            left: 0;
            z-index: 100;
        }
            

        .custom-title {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding-top: 20px;
            padding-bottom: 10px;
            
        }

        .title-text {
            font-size: 1em;
            font-weight: bold;

        }

        .time-box {
            background-color: #85A98F;
            padding: 10px 20px;
            border-radius: 10px;
            font-size: 1em;
            color: #FFFFFF;
        }
    
        .block-container {
            padding-top: 30px !important;
        }
    </style>
    <div class="top-bar"></div>
""", unsafe_allow_html=True)

# ✅ Afficher Titre + Heure
now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

logo_path = "logo-white.png"  # Vérifie que le fichier est bien dans le même dossier

# Lecture du fichier et encodage en base64
logo_base64 = base64.b64encode(Path(logo_path).read_bytes()).decode()

st.markdown(f"""
    <div style="display: flex; align-items: center; justify-content: space-between; background-color: #328E6E; padding: 7px; border-radius: 8px;">
        <div style="flex: 1; display: flex; align-items: center;">
            <img src="data:image/png;base64,{logo_base64}" alt="logo" width="70"  style="margin-top : 10px; margin-right: 18px;"/>
        </div>
        <div style="flex: 2;color:#FFFFFF; text-align: center;margin-top : 10px; font-size: 24px; font-weight: bold;">
            🌐 Axe Chargement digital
        </div>
        <div style="flex: 1;color : White ; text-align: right; font-size: 18px;">
            🕒 {now}
        </div>
    </div>
""", unsafe_allow_html=True)

# Titre centré
# Charger les données depuis le fichier Excel

now = datetime.now()

# Si l'heure est avant 7h, utiliser la feuille du jour précédent
if now.time() < time(7, 0):
    effective_date = now - timedelta(days=1)
else:
    effective_date = now

# Formater la date au format "dd-mm-YYYY"
sheet_name = effective_date.strftime("%d-%m-%Y")

# Charger le fichier Excel et extraire la feuille souhaitée
excel_file = "SituationHFN.xlsx"
df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

# Extraire les données des quais
loading_data = {}

# Parcourir les colonnes de B à AD (index 1 à 30)
for col in range(1, 31):
    quai = df.iloc[6, col]  # Ligne 7 = index 6
    if pd.isna(quai):
        continue

    ship = df.iloc[1, col]             # Ligne 2 = index 1
    quantity_requested = df.iloc[4, col]  # Ligne 5 = index 4
    product_type = df.iloc[5, col]     # Ligne 6 = index 5
    origin = df.iloc[7, col]  
    total_charge = df.iloc[37, 0]  # Ligne 30, colonne B (0-indexed)

    # Cumul de chargement : lignes 13 à 36 (index 12 à 35)
    cumul_data = df.iloc[12:36, col].dropna()

    max_loaded = 0
    tonnage_last_hour = 0

    if not cumul_data.empty:
        max_index = cumul_data.idxmax()  # Ligne où le cumul est max
        max_loaded = cumul_data[max_index]

        # Vérifier que la colonne col+1 existe
        if col + 1 < df.shape[1]:
            tonnage_last_hour = df.iloc[max_index, col + 1]
    
    # Construction de la structure
    if quai not in loading_data:
        loading_data[quai] = {
            "ship": ship,
            "products": {}
        }

    loading_data[quai]["products"][product_type]= {
        "loaded": max_loaded,
        "target": quantity_requested,
        "last_hour": tonnage_last_hour,
        "Source" : origin,
        
    }
print(loading_data)
# Afficher les données dans le tableau de bord
page = st.sidebar.selectbox(
    "Navigation",
    ["Suivi de chargement","CTE"]
)

# Vue 1 : Suivi de chargement
# Vue 1 : Suivi de chargement
# Vue 1 : Suivi de chargement
if page == "Suivi de chargement":
    st.markdown(
    f"""
    <div style='display:flex; justify-content:space-between; align-items:center;'>
        <h3 style="font-size:24px; margin:0;">🚢 Suivi de chargement en temps réel</h3>
        <span style="font-size:20px;color:#FFFFFF; background-color:#F7374F; padding:6px 12px; border-radius:8px;">
            Total Chargé : <strong>{total_charge} t</strong>
        </span>
    </div>
    """,
    unsafe_allow_html=True
)



    # Diviser la page en deux colonnes : Quai 1 (gauche) et Quai 2 (droite)
    col1, col2 = st.columns(2)

    # Afficher les quais de gauche (1NORD, 1BIS, 1TER)
    with col1:
        st.markdown('<h4 style="margin-bottom:10px;">🧭 Quais - 1</h4>', unsafe_allow_html=True)
        for quai, info in loading_data.items():
            if quai.startswith("1"):  # Filtrer les quais de gauche
                # Vérifier si le quai a un navire
                if not info["ship"] or pd.isna(info["ship"]):  # Si pas de navire
                    st.markdown(f'<div style="font-size:18px; font-weight:bold;">Quai {quai} – 🚩 Quai Libre</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div style="font-size:18px; font-weight:bold;">Quai {quai} – Navire : {info["ship"]}</div>', unsafe_allow_html=True)
                    for product, stats in info["products"].items():
                        progress = stats["loaded"] / stats["target"] if stats["target"] > 0 else 0
                        percentage = progress * 100
                        source = stats["Source"]
                        if source == "JFC1":
                            color = "#FFEB3B"  # Jaune
                        elif source == "JFC2":
                           color = "#FFEB3B"  # Gris
                        elif source == "JLN":
                           color = "#FFEB3B"  # Jaune
                        else:
                           color = "#4CAF50"
                        st.markdown(f'<div style="font-size:14px;">🔸 Qualité 🌾 : {product} | Chargé ✅ : {stats["loaded"]} / {stats["target"]} t | Dernière heure 🕒 : {stats["last_hour"]} t </div>', unsafe_allow_html=True)
                        st.progress(progress)
                    
                        st.markdown(f'<div style="font-size:14px; text-align:center;">{percentage:.2f}% Chargé| Source de chargement actuel  🏗️ : {stats["Source"]}</div>', unsafe_allow_html=True)
                        
                st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)

    # Afficher les quais de droite (2NORD, 2SUD, 2BIS, 2TER)
    with col2:
        st.markdown('<h4 style="margin-bottom:10px;">🧭 Quais - 2</h4>', unsafe_allow_html=True)
        for quai, info in loading_data.items():
            if quai.startswith("2"):  # Filtrer les quais de gauche
                # Vérifier si le quai a un navire
                if not info["ship"] or pd.isna(info["ship"]):  # Si pas de navire
                    st.markdown(f'<div style="font-size:18px; font-weight:bold;">Quai {quai} – 🚩 Quai Libre</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div style="font-size:18px; font-weight:bold;">Quai {quai} – Navire : {info["ship"]}</div>', unsafe_allow_html=True)
                    for product, stats in info["products"].items():
                        progress = stats["loaded"] / stats["target"] if stats["target"] > 0 else 0
                        percentage = progress * 100
                        
                        st.markdown(f'<div style="font-size:14px;">🔸 Qualité 🌾 : {product} | Chargé ✅ : {stats["loaded"]} / {stats["target"]} t | Dernière heure 🕒 : {stats["last_hour"]} t </div>', unsafe_allow_html=True)
                        st.progress(progress)
                        st.markdown(f'<div style="font-size:14px; text-align:center;">{percentage:.2f}% Chargé| Source de chargement actuel  🏗️ : {stats["Source"]}</div>', unsafe_allow_html=True)
                st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)


elif page == "CTE":
    from io import BytesIO
    import matplotlib.pyplot as plt
    st.title("Suivi des CTE (Contrôle Tirant d'EAU)")

    with st.form("form_cte"):
        # Récupérer la liste des navires disponibles
        navires_disponibles = []
        for quai, info in loading_data.items():
            if info["ship"] and not pd.isna(info["ship"]):
                navires_disponibles.append(info["ship"])

        navire_choisi = st.selectbox("🚢 Choisir un navire", navires_disponibles, key="navire_choisi")

        # Si le navire sélectionné a changé, réinitialiser la sélection de qualité
        if "prev_navire" not in st.session_state:
            st.session_state.prev_navire = navire_choisi
        elif st.session_state.prev_navire != navire_choisi:
            old_key = f"qualite_choisie_{st.session_state.prev_navire}"
            if old_key in st.session_state:
                del st.session_state[old_key]
            st.session_state.prev_navire = navire_choisi

        # Calculer la liste des qualités associées au navire choisi
        qualites = []
        for quai, info in loading_data.items():
            if info["ship"] == navire_choisi:
                qualites = list(info["products"].keys())
                st.session_state.data_navire = info  # stocker les infos du navire
                break

        # Utiliser une clé dynamique pour la sélection de la qualité
        qualite_choisie = st.selectbox("🌾 Choisir la qualité (produit)", qualites, key=f"qualite_choisie_{navire_choisi}")

        type_cte = st.selectbox("📝 Type de CTE", [
            "Fin de chargement",
            "Changement de Qualité",
            "Changement d'Origine",
            "Vérification de tonnage par le Bord",
            "JPH"
        ])
        # ... (le reste du formulaire)

        valeur_cte = st.number_input("⚖️ Valeur mesurée du CTE (tonnes)", min_value=0.0, format="%.2f")

        submit = st.form_submit_button("✅ Valider")

    if submit:
        # Récupérer les données du suivi à partir des informations stockées
        charge_bascule = st.session_state.data_navire["products"][qualite_choisie]["loaded"]
        target = st.session_state.data_navire["products"][qualite_choisie]["target"]

        ecart = valeur_cte - charge_bascule
        reste_a_charger = target - valeur_cte

        # Création de 2 colonnes : colonne gauche pour le graphique, colonne droite pour les résultats
        col1, col2 = st.columns(2)

        with col1:
            # Graphique
            fig, ax = plt.subplots(figsize=(6, 4))
            ax.bar(
                ["CTE", "Bascule", "Demande"],
                [valeur_cte, charge_bascule, target],
                color=["#1f77b4", "#2ca02c", "#7f7f7f"],
                edgecolor="black",
                linewidth=1.2
            )
            ax.set_ylabel("Tonnage (t)", fontsize=12)
            ax.set_title(f"Comparaison CTE vs Bascule - {navire_choisi}", fontsize=14)
            ax.grid(axis="y", linestyle="--", alpha=0.7)
            st.pyplot(fig)
            buf = BytesIO()
            fig.savefig(buf, format="png")
            buf.seek(0)
            img_data = buf.read()

        with col2:
            # Informations des résultats
            st.markdown(f"### Résultats pour le navire **{navire_choisi}** – Qualité **{qualite_choisie}**")
            st.metric("⚖️ CTE mesuré", f"{valeur_cte} t")
            st.metric("📦 Chargé par bascule", f"{charge_bascule} t")
            st.metric("🔀 Écart CTE vs Bascule", f"{ecart:+.2f} t")
            st.metric("📈 Reste à charger", f"{reste_a_charger:.2f} t")
     
        # Convertir le graphique en image


        # Envoi de l'e-mail
        import os
        from email.message import EmailMessage
        import smtplib
        import streamlit as st
        from dotenv import load_dotenv
        import smtplib
        from email.message import EmailMessage

        # Charger les variables d’environnement
        load_dotenv()

        EMAIL_ADDRESS = "elhafianimouad@gmail.com"
        EMAIL_PASSWORD = "txlg orhh icba bbjf"

        # Remplir dynamiquement les données
        email_destinataire = "mouad.elhafiani.etu21@ensem.ac.ma"
        message = EmailMessage()
        message["Subject"] = f"📩 Rapport CTE – {navire_choisi}"
        message["From"] = EMAIL_ADDRESS
        message["To"] = email_destinataire

        message.set_content(f"""
        Rapport CTE
        Navire : {navire_choisi}
        Produit : {qualite_choisie}
        Type de CTE : {type_cte}

        📊 Résumé :
        - Valeur CTE : {valeur_cte} t
        - Chargé par bascule : {charge_bascule} t
        - Écart : {ecart:.2f} t
        - Reste à charger : {reste_a_charger:.2f} t
        """)
        html_content = f"""
        <html>
            <body>
                <p>
                    Navire : {navire_choisi}<br>
                    Produit : {qualite_choisie}<br>
                    Type de CTE : {type_cte}<br><br>
                    📊 Résumé :<br>
                    - Valeur CTE : {valeur_cte} t<br>
                    - Chargé par bascule : {charge_bascule} t<br>
                    - Écart : {ecart:.2f} t<br>
                    - Reste à charger : {reste_a_charger:.2f} t
                </p>      
            <p>Image intégrée dans le contenu :</p>
            <p><img src="cid:graphique"></p>
        </body>
        </html>
        """
        message.add_alternative(html_content, subtype="html")
        message.get_payload()[1].add_related(img_data, maintype="image", subtype="png", cid="<graphique>")

    
        try:
            with smtplib.SMTP("smtp.gmail.com", 587, timeout=20) as server:
                server.starttls()
                server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)  # Pour Gmail, assurez-vous d'utiliser un mot de passe d'application si 2FA est activé
                server.send_message(message)
            st.success("✉️ E-mail envoyé avec succès !")
        except Exception as e:
            st.error(f"Erreur lors de l'envoi de l'email : {e}")


        # Enregistrer l'événement CTE dans l'historique (les 10 derniers)
        cte_record = {
            "Navire": navire_choisi,
            "Qualité": qualite_choisie,
            "Type de CTE": type_cte,
            "CTE mesuré (t)": valeur_cte,
            "Bascule (t)": charge_bascule,
            "Écart (t)": round(ecart, 2),
            "Reste à charger (t)": round(reste_a_charger, 2),
            "Horodatage": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        }
        if "cte_history" not in st.session_state:
            st.session_state["cte_history"] = []
        st.session_state["cte_history"].append(cte_record)
        # Conserver uniquement les 10 derniers enregistrements
        st.session_state["cte_history"] = st.session_state["cte_history"][-10:]
        # Inverser l'ordre pour afficher le dernier en premier
        cte_history_display = list(reversed(st.session_state["cte_history"]))

        st.markdown("### Dernières 10 CTE effectuées")
        st.table(pd.DataFrame(cte_history_display))
        

