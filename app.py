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

cal, formateurs, ecues, saisies = {}, {}, {}, {}
mode_api_ok = False

# Tentative de connexion directe à l'API Grist
if "GRIST_API_KEY" in st.secrets:
    try:
        headers = {"Authorization": f"Bearer {st.secrets['GRIST_API_KEY'].strip()}"}
        r_cal = requests.get(f"{GRIST_URL}/tables/Calendrier_Cadre/records", headers=headers, timeout=8)
        if r_cal.status_code == 200:
            cal = {r["id"]: r["fields"] for r in r_cal.json().get("records", [])}
            formateurs = {r["id"]: r["fields"] for r in requests.get(f"{GRIST_URL}/tables/Formateurs/records", headers=headers, timeout=8).json().get("records", [])}
            ecues = {r["id"]: r["fields"] for r in requests.get(f"{GRIST_URL}/tables/Maquette_ECUE/records", headers=headers, timeout=8).json().get("records", [])}
            saisies = {r["id"]: r["fields"] for r in requests.get(f"{GRIST_URL}/tables/Saisies_EDT/records", headers=headers, timeout=8).json().get("records", [])}
            mode_api_ok = True
    except Exception:
        mode_api_ok = False

if not mode_api_ok:
    st.info("🔒 **Mode hors-ligne (Pare-feu État) :** Le serveur Grist bloque les connexions directes depuis Streamlit Cloud. Dans Grist, cliquez sur **Outils > Télécharger le document (.xlsx)** et déposez le fichier ci-dessous :")
    fichier_xlsx = st.file_uploader("Importer l'export Excel (.xlsx) de votre document Grist :", type=["xlsx"])
    if not fichier_xlsx:
        st.stop()
    xls = pd.ExcelFile(fichier_xlsx)
    df_c = pd.read_excel(xls, "Calendrier_Cadre").fillna("")
    df_f = pd.read_excel(xls, "Formateurs").fillna("")
    df_e = pd.read_excel(xls, "Maquette_ECUE").fillna("")
    df_s = pd.read_excel(xls, "Saisies_EDT").fillna("")

    cal = {i + 1: row.to_dict() for i, row in df_c.iterrows()}
    formateurs = {i + 1: row.to_dict() for i, row in df_f.iterrows()}
    ecues = {i + 1: row.to_dict() for i, row in df_e.iterrows()}
    saisies = {i + 1: row.to_dict() for i, row in df_s.iterrows()}
else:
    if st.sidebar.button("🔄 Actualiser depuis Grist", use_container_width=True):
        st.rerun()

disciplines = sorted(list({str(f.get("Discipline", "")) for f in ecues.values() if f.get("Discipline")}))
choix_disc = st.sidebar.selectbox("Filtrer par discipline :", ["Toutes les disciplines"] + disciplines)

onglet1, onglet2, onglet3 = st.tabs(["🗓️ Emploi du temps & Export Word", "⏳ Suivi des heures & Relances groupées", "⚠️ Contrôle des conflits"])

with onglet1:
    lignes_edt = []
    for sid, s in saisies.items():
        ecue_val = s.get("ECUE")
        ecue_obj = ecues.get(ecue_val, {}) if isinstance(ecue_val, int) else next((v for v in ecues.values() if v.get("Code_ECUE") == ecue_val), {})
        disc = str(ecue_obj.get("Discipline", "Inconnue"))
        if choix_disc != "Toutes les disciplines" and disc != choix_disc:
            continue

        date_val = s.get("Date_Seance")
        date_obj = cal.get(date_val, {}) if isinstance(date_val, int) else next((v for v in cal.values() if str(date_val) in str(v.get("Libelle_Date", "")) or str(date_val) in str(v.get("Date_Jour", ""))), {})
        ts = date_obj.get("Date_Jour")
        if isinstance(ts, (int, float)):
            date_str = datetime.fromtimestamp(ts).strftime("%d/%m/%Y")
            tri = ts
        else:
            date_str = str(ts)[:10] if ts else "Date non définie"
            tri = 0

        form_val = s.get("Intervenant")
        nom_form = formateurs.get(form_val, {}).get("Nom_Complet", str(form_val)) if isinstance(form_val, int) else str(form_val)

        lignes_edt.append({
            "Tri_Date": tri,
            "Semaine": date_obj.get("Semaine", ""),
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

        doc = Document()
        doc.add_heading(f"Emploi du Temps Prévisionnel — {choix_disc}", level=1)
        doc.add_paragraph(f"Document généré le {datetime.now().strftime('%d/%m/%Y à %H:%M')} depuis la base Grist INSPE / Rectorat.")

        table_w = doc.add_table(rows=1, cols=7)
        table_w.style = "Table Grid"
        for i, h in enumerate(["Sem.", "Date", "Créneau", "ECUE", "Intervenant", "Durée", "Salle / Besoins"]):
            table_w.rows[0].cells[i].text = h

        for _, row in df_edt.iterrows():
            cells = table_w.add_row().cells
            for idx, col_name in enumerate(["Semaine", "Date", "Créneau", "ECUE", "Intervenant", "Durée (h)", "Salle / Matériel"]):
                cells[idx].text = f"{row[col_name]} h" if col_name == "Durée (h)" else str(row[col_name])

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button("📘 Télécharger le planning officiel au format Word (.docx)", data=buf, file_name=f"EDT_{choix_disc.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
    else:
        st.info("Aucune séance planifiée pour cette sélection.")

with onglet2:
    st.subheader("Avancement des saisies par ECUE et relances")
    a_relancer, tableau_suivi = [], []
    for eid, e in ecues.items():
        disc = str(e.get("Discipline", ""))
        if choix_disc != "Toutes les disciplines" and disc != choix_disc:
            continue
        h_maq = float(e.get("Heures_Maquette") or 0)
        h_plan = float(e.get("Heures_Planifiees") or 0)
        reste = h_maq - h_plan
        resp_val = e.get("Responsable")
        resp_obj = formateurs.get(resp_val, {}) if isinstance(resp_val, int) else next((v for v in formateurs.values() if v.get("Nom_Complet") == resp_val), {})
        nom_resp = resp_obj.get("Nom_Complet", str(resp_val))
        mail_resp = resp_obj.get("Email", "")

        tableau_suivi.append({"Discipline": disc, "Code ECUE": e.get("Code_ECUE", ""), "Intitulé": e.get("Intitule", ""), "Responsable": nom_resp, "Maquette (h)": h_maq, "Planifié (h)": h_plan, "Reste (h)": reste, "État": e.get("Statut_Suivi", "")})
        if reste > 0:
            a_relancer.append({"email": mail_resp, "nom": nom_resp, "code": e.get("Code_ECUE", ""), "intitule": e.get("Intitule", ""), "reste": reste})

    if tableau_suivi:
        st.dataframe(pd.DataFrame(tableau_suivi), use_container_width=True, hide_index=True)
    st.divider()
    if a_relancer:
        st.warning(f"⏳ **{len(a_relancer)} ECUE incomplets** nécessitent une relance.")
        st.text_input("📋 Liste des courriels (champ Cci) :", value="; ".join(sorted(list({r["email"] for r in a_relancer if r["email"]}))))
        txt = "Bonjour,\n\nDes créneaux restent à positionner sur les enseignements suivants :\n\n" + "\n".join([f"- {r['code']} ({r['intitule']}) — Responsable : {r['nom']} -> Reste : {r['reste']} h" for r in a_relancer]) + "\n\nMerci de compléter vos vœux via le formulaire Grist.\n\nCordialement"
        st.text_area("✉️ Message type de relance groupée :", value=txt, height=200)
    else:
        st.success("✅ Tous les ECUE de cette sélection sont complets !")

with onglet3:
    conflits = [l for l in lignes_edt if l["Statut"] != "✅ OK"]
    if conflits:
        st.error(f"🚨 {len(conflits)} anomalie(s) détectée(s) :")
        st.dataframe(pd.DataFrame(conflits), use_container_width=True, hide_index=True)
    else:
        st.success("✅ Aucun chevauchement de créneau ni conflit détecté.")
