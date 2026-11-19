import streamlit as st
import requests
import pandas as pd
from datetime import datetime
from docx import Document
import io

st.set_page_config(page_title="Pilotage EDT Formation - Grist", page_icon="📅", layout="wide")
st.title("📅 Pilotage des Emplois du Temps — Formation Initiale (INSPE / Rectorat)")

DOC_ID = "3pz5iXJkTkwR"
GRIST_URL = f"https://numerique.gouv.fr{DOC_ID}"

def lire_table(nom_table):
    api_key = st.secrets["GRIST_API_KEY"].strip()
    rep = requests.get(f"{GRIST_URL}/tables/{nom_table}/records", headers={"Authorization": f"Bearer {api_key}"}, timeout=30)
    if rep.status_code == 200:
        return {r["id"]: r["fields"] for r in rep.json().get("records", [])}
    st.error(f"Erreur de lecture Grist sur {nom_table} ({rep.status_code}) : {rep.text}")
    return {}

if st.sidebar.button("🔄 Actualiser depuis Grist", use_container_width=True):
    st.rerun()

# Chargement des 4 tables liées
cal = lire_table("Calendrier_Cadre")
formateurs = lire_table("Formateurs")
ecues = lire_table("Maquette_ECUE")
saisies = lire_table("Saisies_EDT")

# Sélecteur de discipline dans la barre latérale
disciplines = sorted(list({f.get("Discipline", "Non précisée") for f in ecues.values() if f.get("Discipline")}))
choix_disc = st.sidebar.selectbox("Filtrer par discipline :", ["Toutes les disciplines"] + disciplines)

onglet1, onglet2, onglet3 = st.tabs(["🗓️ Emploi du temps & Export Word", "⏳ Suivi des heures & Relances groupées", "⚠️ Contrôle des conflits"])

# ==============================================================================
# ONGLET 1 : PLANNING CONSOLIDÉ & EXPORT WORD
# ==============================================================================
with onglet1:
    lignes_edt = []
    for sid, s in saisies.items():
        id_ecue = s.get("ECUE")
        ecue_obj = ecues.get(id_ecue, {})
        disc = ecue_obj.get("Discipline", "Inconnue")
        if choix_disc != "Toutes les disciplines" and disc != choix_disc:
            continue

        id_date = s.get("Date_Seance")
        date_obj = cal.get(id_date, {})
        ts = date_obj.get("Date_Jour")
        date_str = datetime.fromtimestamp(ts).strftime("%d/%m/%Y") if isinstance(ts, (int, float)) else "Date non définie"
        semaine = date_obj.get("Semaine", "")

        id_form = s.get("Intervenant")
        nom_form = formateurs.get(id_form, {}).get("Nom_Complet", "Non attribué")

        lignes_edt.append({
            "Tri_Date": ts if isinstance(ts, (int, float)) else 0,
            "Semaine": semaine,
            "Date": date_str,
            "Créneau": s.get("Creneau", ""),
            "Discipline": disc,
            "ECUE": f"{ecue_obj.get('Code_ECUE', '')} - {ecue_obj.get('Intitule', '')}",
            "Intervenant": nom_form,
            "Durée (h)": s.get("Duree_Heures", 0),
            "Salle / Matériel": s.get("Salle_Besoins", ""),
            "Statut": s.get("Alerte_Conflit", "")
        })

    if lignes_edt:
        df_edt = pd.DataFrame(lignes_edt).sort_values(by=["Tri_Date", "Créneau"]).drop(columns=["Tri_Date"])
        st.dataframe(df_edt, use_container_width=True, hide_index=True)

        # Génération de l'export Word mis en page
        doc = Document()
        doc.add_heading(f"Emploi du Temps Prévisionnel — {choix_disc}", level=1)
        doc.add_paragraph(f"Document généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')} depuis la base Grist INSPE / Rectorat.")

        table_w = doc.add_table(rows=1, cols=7)
        table_w.style = "Table Grid"
        entetes = ["Sem.", "Date", "Créneau", "ECUE", "Intervenant", "Durée", "Salle / Besoins"]
        for i, h in enumerate(entetes):
            table_w.rows[0].cells[i].text = h

        for _, row in df_edt.iterrows():
            cells = table_w.add_row().cells
            cells[0].text = str(row["Semaine"])
            cells[1].text = str(row["Date"])
            cells[2].text = str(row["Créneau"])
            cells[3].text = str(row["ECUE"])
            cells[4].text = str(row["Intervenant"])
            cells[5].text = f"{row['Durée (h)']} h"
            cells[6].text = str(row["Salle / Matériel"])

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button(
            "📘 Télécharger le planning officiel au format Word (.docx)",
            data=buf,
            file_name=f"EDT_{choix_disc.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    else:
        st.info("Aucune séance planifiée pour cette sélection.")

# ==============================================================================
# ONGLET 2 : SUIVI MAQUETTE & RELANCES GROUPÉES
# ==============================================================================
with onglet2:
    st.subheader("Avancement des saisies par ECUE et relances")
    a_relancer = []
    tableau_suivi = []

    for eid, e in ecues.items():
        disc = e.get("Discipline", "")
        if choix_disc != "Toutes les disciplines" and disc != choix_disc:
            continue

        h_maq = e.get("Heures_Maquette", 0) or 0
        h_plan = e.get("Heures_Planifiees", 0) or 0
        reste = h_maq - h_plan
        resp_id = e.get("Responsable")
        resp_obj = formateurs.get(resp_id, {})
        nom_resp = resp_obj.get("Nom_Complet", "Non défini")
        mail_resp = resp_obj.get("Email", "")

        tableau_suivi.append({
            "Discipline": disc,
            "Code ECUE": e.get("Code_ECUE", ""),
            "Intitulé": e.get("Intitule", ""),
            "Responsable": nom_resp,
            "Maquette (h)": h_maq,
            "Planifié (h)": h_plan,
            "Reste (h)": reste,
            "État": e.get("Statut_Suivi", "")
        })

        if reste > 0:
            a_relancer.append({
                "email": mail_resp,
                "nom": nom_resp,
                "code": e.get("Code_ECUE", ""),
                "intitule": e.get("Intitule", ""),
                "reste": reste
            })

    if tableau_suivi:
        st.dataframe(pd.DataFrame(tableau_suivi), use_container_width=True, hide_index=True)

    st.divider()
    if a_relancer:
        st.warning(f"⏳ **{len(a_relancer)} ECUE incomplets** nécessitent une relance.")
        emails_groupes = "; ".join(sorted(list({r["email"] for r in a_relancer if r["email"]})))
        st.text_input("📋 Liste des courriels à copier dans votre messagerie (champ Cci) :", value=emails_groupes)

        texte_relance = "Bonjour,\n\nDans le cadre de la construction des emplois du temps de formation initiale (INSPE / Rectorat), des créneaux restent à positionner sur les enseignements suivants :\n\n"
        for r in a_relancer:
            texte_relance += f"- {r['code']} ({r['intitule']}) — Responsable : {r['nom']} -> Reste à planifier : {r['reste']} h\n"
        texte_relance += "\nMerci de compléter vos vœux via le formulaire Grist dédié dès que possible.\n\nCordialement,\nL'équipe de pilotage"

        st.text_area("✉️ Message type de relance groupée prêt à copier :", value=texte_relance, height=220)
    else:
        st.success("✅ Tous les ECUE de cette sélection sont complets !")

# ==============================================================================
# ONGLET 3 : CONTRÔLE DES CONFLITS ET ANOMALIES
# ==============================================================================
with onglet3:
    conflits = [l for l in lignes_edt if l["Statut"] != "✅ OK"]
    if conflits:
        st.error(f"🚨 {len(conflits)} anomalie(s) détectée(s) dans les saisies :")
        st.dataframe(pd.DataFrame(conflits), use_container_width=True, hide_index=True)
    else:
        st.success("✅ Aucun chevauchement de créneau ni conflit avec les journées banalisées Rectorat.")

