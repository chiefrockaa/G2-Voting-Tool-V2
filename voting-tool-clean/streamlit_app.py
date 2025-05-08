import streamlit as st
import pandas as pd
import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials
from collections import defaultdict
from io import BytesIO

st.set_page_config(page_title="G2 Voting Tool", layout="centered")
st.title("🔝 G2 Voting Tool V4 – mit Google Sheets Speicherung")

# ✅ Sicherer Zugriff auf Service Account Secret
raw_key = st.secrets["gcp_service_account"]

if isinstance(raw_key, str):
    cleaned = raw_key.encode().decode("unicode_escape")
    service_json = json.loads(cleaned)
else:
    service_json = raw_key

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_json, scope)
client = gspread.authorize(creds)

# 🔧 Name deines Google Sheets (vorher erstellen!)
SHEET_NAME = "G2_Votings_DB"

# Liste vorhandener Tabs (je Voting ein Tab)
tabs = client.open(SHEET_NAME).worksheets()
tab_names = [t.title for t in tabs]

st.sidebar.header("🗳️ Voting auswählen oder erstellen")
voting_wahl = st.sidebar.selectbox("Wähle ein Voting", tab_names + ["Neues Voting erstellen"])

# Neues Voting-Tab anlegen
if voting_wahl == "Neues Voting erstellen":
    neuer_voting_name = st.sidebar.text_input("Name für neues Voting")
    if st.sidebar.button("Erstellen"):
        if neuer_voting_name.strip():
            client.open(SHEET_NAME).add_worksheet(title=neuer_voting_name.strip(), rows="100", cols="12")
            st.success(f"Voting '{neuer_voting_name}' wurde erstellt. Lade die Seite neu.")
            st.stop()
        else:
            st.sidebar.warning("Bitte gib einen Namen ein.")
    st.stop()
else:
    sheet = client.open(SHEET_NAME).worksheet(voting_wahl)

# Verwaltung
st.sidebar.subheader("🧹 Verwaltung")
if st.sidebar.checkbox("⚠️ Ich will dieses Voting wirklich zurücksetzen"):
    if st.sidebar.button("Voting zurücksetzen (leeren)"):
        sheet.clear()
        st.sidebar.success(f"Voting '{voting_wahl}' wurde geleert.")
        st.stop()

# Eingabe
st.header(f"📋 Voting: {voting_wahl}")
name = st.text_input("Dein Name")
spiele = [st.text_input(f"Platz {i+1} ({10 - i} Punkte)", key=i) for i in range(10)]

if st.button("Einreichen"):
    if not name.strip():
        st.warning("Bitte gib deinen Namen ein.")
    elif not any(spiele):
        st.warning("Bitte gib mindestens ein Spiel an.")
    else:
        sheet.append_row([name] + spiele)
        st.success("✅ Stimme gespeichert!")

# Gesamtranking
if st.checkbox("📊 Gesamtranking anzeigen"):
    rows = sheet.get_all_values()
    if len(rows) < 1:
        st.info("Noch keine Daten vorhanden.")
    else:
        df = pd.DataFrame(rows)
        df.columns = [f"Platz {i}" if i > 0 else "Name" for i in range(len(df.columns))]
        spiele_punkte = defaultdict(int)
        spiele_quellen = defaultdict(list)

        for _, row in df.iterrows():
            voter = row[0]
            for i, spiel in enumerate(row[1:]):
                if spiel.strip():
                    punkte = 10 - i
                    spiele_punkte[spiel.strip()] += punkte
                    spiele_quellen[spiel.strip()].append(f"{voter} ({punkte} P)")

        ranking = sorted(spiele_punkte.items(), key=lambda x: x[1], reverse=True)
        ranking_df = pd.DataFrame([
            {
                "Spiel": spiel,
                "Gesamtpunkte": punkte,
                "Beitragsquellen": ", ".join(spiele_quellen[spiel])
            }
            for spiel, punkte in ranking
        ])

        st.subheader("🏆 Gesamtranking")
        st.dataframe(ranking_df, use_container_width=True)

        # Excel-Export
        st.markdown("### ⬇️ Ranking exportieren")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            ranking_df.to_excel(writer, index=False, sheet_name="Gesamtranking")
        st.download_button("📥 Excel herunterladen", data=output.getvalue(), file_name=f"{voting_wahl}_ranking.xlsx")
