import streamlit as st
import pandas as pd
from collections import defaultdict
import os
from io import BytesIO

st.title("🔝 Top-10 Voting Tool")

DATA_DIR = "votings"
os.makedirs(DATA_DIR, exist_ok=True)

# Bestehende Votings
bestehende_votings = [f.replace(".csv", "") for f in os.listdir(DATA_DIR) if f.endswith(".csv")]

# Auswahl oder neues Voting
st.sidebar.header("🗳️ Voting auswählen oder erstellen")
voting_wahl = st.sidebar.selectbox("Wähle ein Voting", bestehende_votings + ["Neues Voting erstellen"])

# Neues Voting erstellen
if voting_wahl == "Neues Voting erstellen":
    neuer_voting_name = st.sidebar.text_input("Name für neues Voting")
    if st.sidebar.button("Erstellen"):
        if neuer_voting_name.strip():
            new_path = os.path.join(DATA_DIR, f"{neuer_voting_name.strip()}.csv")
            open(new_path, "a").close()
            st.success(f"Voting '{neuer_voting_name}' wurde erstellt. Starte neu, um es auszuwählen.")
            st.stop()
        else:
            st.sidebar.warning("Bitte gib einen Namen ein.")
    st.stop()
else:
    voting_name = voting_wahl
    csv_path = os.path.join(DATA_DIR, f"{voting_name}.csv")

# Verwaltung: zurücksetzen oder löschen
st.sidebar.subheader("🧹 Verwaltung")
if st.sidebar.button("Voting zurücksetzen (leeren)"):
    open(csv_path, "w").close()
    st.sidebar.success(f"Inhalte von '{voting_name}' wurden gelöscht.")
    st.stop()

if st.sidebar.button("Voting komplett löschen"):
    os.remove(csv_path)
    st.sidebar.success(f"Voting '{voting_name}' wurde gelöscht. Starte neu.")
    st.stop()

# Eingabeformular
st.header(f"📋 Voting: {voting_name}")
name = st.text_input("Dein Name")
spiele = [st.text_input(f"Platz {i+1} ({10 - i} Punkte)", key=i) for i in range(10)]

if st.button("Einreichen"):
    if not name.strip():
        st.warning("Bitte gib deinen Namen ein.")
    elif not any(spiele):
        st.warning("Bitte gib mindestens ein Spiel an.")
    else:
        with open(csv_path, "a") as f:
            f.write(name + "," + ",".join(spiele) + "\n")
        st.success("Danke! Deine Stimme wurde gespeichert.")

# Gesamtranking anzeigen
if st.checkbox("📊 Gesamtranking anzeigen"):
    if not os.path.exists(csv_path) or os.stat(csv_path).st_size == 0:
        st.info("Noch keine Daten vorhanden.")
    else:
        df = pd.read_csv(csv_path, header=None)
        spiele_punkte = defaultdict(int)
        spiele_quellen = defaultdict(list)

        for _, row in df.iterrows():
            voter = row[0]
            for i, spiel in enumerate(row[1:]):
                if pd.notna(spiel) and spiel.strip():
                    spiel_clean = spiel.strip()
                    punkte = 10 - i
                    spiele_punkte[spiel_clean] += punkte
                    spiele_quellen[spiel_clean].append(f"{voter} ({punkte} P)")

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

        # Exportfunktion
        st.markdown("### ⬇️ Ranking exportieren")
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            ranking_df.to_excel(writer, index=False, sheet_name="Gesamtranking")
        st.download_button("📥 Excel herunterladen", data=output.getvalue(), file_name=f"{voting_name}_ranking.xlsx")
