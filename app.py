import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
import re

st.set_page_config(
    page_title="Confronto fasce e anomalie",
    layout="wide"
)

st.title("Confronto fasce e anomalie")
st.write(
    "App per confrontare due file Excel e rilevare possibili sovrapposizioni "
    "tra fasce/orari riferite alla stessa matricola e alla stessa data."
)

# =========================
# FUNZIONI DI SUPPORTO
# =========================

def carica_excel(file):
    try:
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"Errore nella lettura del file: {e}")
        return None


def normalizza_matricola(valore):
    if pd.isna(valore):
        return None

    testo = str(valore).strip()

    if testo.lower() in ["none", "nan", ""]:
        return None

    # Se arriva come 141212.0 lo trasformo in 141212
    if testo.endswith(".0"):
        testo = testo[:-2]

    return testo


def normalizza_data(valore):
    if pd.isna(valore):
        return None

    try:
        return pd.to_datetime(valore).date()
    except Exception:
        return None


def crea_data_da_giorno(anno, mese, giorno):
    try:
        giorno_int = int(giorno)
        return datetime(anno, mese, giorno_int).date()
    except Exception:
        return None


def parse_ora(testo):
    """
    Converte stringhe tipo:
    8
    08
    8:00
    08:00
    in minuti dalla mezzanotte.
    """
    if testo is None:
        return None

    testo = str(testo).strip()

    if ":" in testo:
        pezzi = testo.split(":")
        try:
            ore = int(pezzi[0])
            minuti = int(pezzi[1])
            return ore * 60 + minuti
        except Exception:
            return None

    try:
        ore = int(testo)
        return ore * 60
    except Exception:
        return None


def estrai_intervalli_da_testo(testo):
    """
    Cerca intervalli orari in testi tipo:
    8-20 / 20-8 DIRIGENZA
    14 - 20 / 20 - 8 DIRIGENZA
    08:00-14:00
    20:00 - 08:00
    """

    if pd.isna(testo):
        return []

    testo = str(testo)

    if testo.lower() in ["none", "nan", ""]:
        return []

    pattern = r"(\d{1,2}(?::\d{2})?)\s*-\s*(\d{1,2}(?::\d{2})?)"
    matches = re.findall(pattern, testo)

    intervalli = []

    for inizio_txt, fine_txt in matches:
        inizio = parse_ora(inizio_txt)
        fine = parse_ora(fine_txt)

        if inizio is None or fine is None:
            continue

        # Se la fine è minore o uguale all'inizio, considero turno a cavallo della mezzanotte
        if fine <= inizio:
            fine = fine + 24 * 60

        intervalli.append((inizio, fine))

    return intervalli


def minuti_to_ora(minuti):
    """
    Converte minuti in formato HH:MM.
    Gestisce anche orari oltre la mezzanotte.
    """
    minuti = int(minuti)
    minuti_mod = minuti % (24 * 60)
    ore = minuti_mod // 60
    mins = minuti_mod % 60
    return f"{ore:02d}:{mins:02d}"


def calcola_sovrapposizione(intervalli_a, intervalli_b):
    """
    Restituisce lista di sovrapposizioni tra due liste di intervalli.
    Ogni intervallo è espresso in minuti.
    """
    risultati = []

    for a_start, a_end in intervalli_a:
        for b_start, b_end in intervalli_b:
            start = max(a_start, b_start)
            end = min(a_end, b_end)

            if start < end:
                risultati.append({
                    "inizio_sovrapposizione": minuti_to_ora(start),
                    "fine_sovrapposizione": minuti_to_ora(end),
                    "minuti_sovrapposizione": end - start
                })

    return risultati


def trasforma_file2(df_b, anno, mese):
    """
    Trasforma il file 2 da formato largo a formato lungo.

    Da:
    Cal | Orario 1 | Decodifica Orario 1 | Matricola 1 | ...

    A:
    Data | Progressivo | Orario | Decodifica Orario | Matricola
    """

    righe = []

    for _, row in df_b.iterrows():
        giorno = row.get("Cal")
        data = crea_data_da_giorno(anno, mese, giorno)

        if data is None:
            continue

        for i in range(1, 11):
            col_orario = f"Orario {i}"
            col_dec_orario = f"Decodifica di Orario {i}"
            col_matricola = f"Matricola {i}"
            col_dec_matricola = f"Decodifica di Matricola {i}"

            if col_orario not in df_b.columns:
                continue

            orario = row.get(col_orario)
            dec_orario = row.get(col_dec_orario) if col_dec_orario in df_b.columns else None
            matricola = row.get(col_matricola) if col_matricola in df_b.columns else None
            dec_matricola = row.get(col_dec_matricola) if col_dec_matricola in df_b.columns else None

            matricola_norm = normalizza_matricola(matricola)

            if matricola_norm is None:
                continue

            righe.append({
                "Data": data,
                "Progressivo": i,
                "Orario": orario,
                "Decodifica Orario": dec_orario,
                "Matricola": matricola_norm,
                "Decodifica Matricola": dec_matricola,
                "Intervalli": estrai_intervalli_da_testo(dec_orario)
            })

    return pd.DataFrame(righe)


def prepara_file1(df_a, colonna_fascia_file1):
    """
    Normalizza il file 1.
    """

    df = df_a.copy()

    df["Matricola_norm"] = df["Matricola"].apply(normalizza_matricola)
    df["Data_norm"] = df["Data Rif."].apply(normalizza_data)
    df["Fascia_file1"] = df[colonna_fascia_file1].astype(str)
    df["Intervalli_file1"] = df["Fascia_file1"].apply(estrai_intervalli_da_testo)

    return df


def genera_report(df_a_norm, df_b_long):
    anomalie = []

    for _, riga_a in df_a_norm.iterrows():
        matricola_a = riga_a.get("Matricola_norm")
        data_a = riga_a.get("Data_norm")

        if matricola_a is None or data_a is None:
            continue

        corrispondenze_b = df_b_long[
            (df_b_long["Matricola"] == matricola_a) &
            (df_b_long["Data"] == data_a)
        ]

        if corrispondenze_b.empty:
            continue

        for _, riga_b in corrispondenze_b.iterrows():
            intervalli_a = riga_a.get("Intervalli_file1", [])
            intervalli_b = riga_b.get("Intervalli", [])

            sovrapposizioni = calcola_sovrapposizione(intervalli_a, intervalli_b)

            if sovrapposizioni:
                for sov in sovrapposizioni:
                    anomalie.append({
                        "Tipo anomalia": "Sovrapposizione oraria",
                        "Matricola": matricola_a,
                        "Cognome": riga_a.get("Cognome"),
                        "Nome": riga_a.get("Nome"),
                        "Data": data_a,
                        "Fascia file 1": riga_a.get("Fascia_file1"),
                        "Orario file 2": riga_b.get("Orario"),
                        "Fascia file 2": riga_b.get("Decodifica Orario"),
                        "Progressivo file 2": riga_b.get("Progressivo"),
                        "Inizio sovrapposizione": sov["inizio_sovrapposizione"],
                        "Fine sovrapposizione": sov["fine_sovrapposizione"],
                        "Minuti sovrapposizione": sov["minuti_sovrapposizione"],
                        "Nota": "Stessa matricola, stessa data, fasce orarie sovrapposte"
                    })
            else:
                anomalie.append({
                    "Tipo anomalia": "Coincidenza matricola/data senza orario interpretabile",
                    "Matricola": matricola_a,
                    "Cognome": riga_a.get("Cognome"),
                    "Nome": riga_a.get("Nome"),
                    "Data": data_a,
                    "Fascia file 1": riga_a.get("Fascia_file1"),
                    "Orario file 2": riga_b.get("Orario"),
                    "Fascia file 2": riga_b.get("Decodifica Orario"),
                    "Progressivo file 2": riga_b.get("Progressivo"),
                    "Inizio sovrapposizione": "",
                    "Fine sovrapposizione": "",
                    "Minuti sovrapposizione": "",
                    "Nota": "Presente nello stesso giorno nei due file, ma una delle due fasce non contiene un intervallo orario leggibile"
                })

    return pd.DataFrame(anomalie)


def crea_excel_report(df_report, df_file2_trasformato):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_report.to_excel(writer, index=False, sheet_name="Anomalie")
        df_file2_trasformato.drop(columns=["Intervalli"], errors="ignore").to_excel(
            writer,
            index=False,
            sheet_name="File2 trasformato"
        )

        workbook = writer.book

        formato_header = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAF7",
            "border": 1
        })

        formato_bordo = workbook.add_format({
            "border": 1
        })

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)

            if sheet_name == "Anomalie":
                df = df_report
            else:
                df = df_file2_trasformato.drop(columns=["Intervalli"], errors="ignore")

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, formato_header)
                worksheet.set_column(col_num, col_num, 22)

            for row_num in range(1, len(df) + 1):
                for col_num in range(len(df.columns)):
                    worksheet.write(row_num, col_num, df.iloc[row_num - 1, col_num], formato_bordo)

    output.seek(0)
    return output


# =========================
# INTERFACCIA APP
# =========================

st.subheader("1. Caricamento file")

file_a = st.file_uploader("Carica il primo file Excel", type=["xlsx"], key="file_a")
file_b = st.file_uploader("Carica il secondo file Excel", type=["xlsx"], key="file_b")

st.subheader("2. Parametri del confronto")

col_mese, col_anno = st.columns(2)

with col_mese:
    mese = st.selectbox(
        "Mese di riferimento del file 2",
        options=list(range(1, 13)),
        index=3,
        format_func=lambda x: {
            1: "Gennaio",
            2: "Febbraio",
            3: "Marzo",
            4: "Aprile",
            5: "Maggio",
            6: "Giugno",
            7: "Luglio",
            8: "Agosto",
            9: "Settembre",
            10: "Ottobre",
            11: "Novembre",
            12: "Dicembre"
        }[x]
    )

with col_anno:
    anno = st.number_input(
        "Anno di riferimento del file 2",
        min_value=2020,
        max_value=2035,
        value=2026,
        step=1
    )

if file_a is not None and file_b is not None:
    df_a = carica_excel(file_a)
    df_b = carica_excel(file_b)

    if df_a is not None and df_b is not None:

        st.subheader("3. Anteprima dei file")

        col1, col2 = st.columns(2)

        with col1:
            st.write("Anteprima file 1")
            st.dataframe(df_a.head(20), use_container_width=True)

        with col2:
            st.write("Anteprima file 2")
            st.dataframe(df_b.head(20), use_container_width=True)

        st.subheader("4. Scelta della colonna fascia/orario del file 1")

        colonne_file1 = list(df_a.columns)

        default_colonna = "Desc. Or.PD" if "Desc. Or.PD" in colonne_file1 else colonne_file1[0]

        colonna_fascia_file1 = st.selectbox(
            "Quale colonna del file 1 contiene la fascia/orario da confrontare?",
            options=colonne_file1,
            index=colonne_file1.index(default_colonna)
        )

        st.info(
            "Nel file 2 l'app usa automaticamente le colonne ripetute "
            "Orario 1-10, Decodifica di Orario 1-10 e Matricola 1-10."
        )

        if st.button("Avvia confronto"):

            colonne_obbligatorie_file1 = ["Matricola", "Cognome", "Nome", "Data Rif."]
            colonne_mancanti = [c for c in colonne_obbligatorie_file1 if c not in df_a.columns]

            if colonne_mancanti:
                st.error(f"Nel file 1 mancano queste colonne obbligatorie: {colonne_mancanti}")
            elif "Cal" not in df_b.columns:
                st.error("Nel file 2 manca la colonna obbligatoria 'Cal'.")
            else:
                df_a_norm = prepara_file1(df_a, colonna_fascia_file1)
                df_b_long = trasforma_file2(df_b, int(anno), int(mese))

                st.subheader("5. File 2 trasformato")
                st.write(
                    "Questa tabella mostra il file 2 trasformato in formato confrontabile: "
                    "una riga per ogni matricola/fascia/data."
                )
                st.dataframe(
                    df_b_long.drop(columns=["Intervalli"], errors="ignore"),
                    use_container_width=True
                )

                report = genera_report(df_a_norm, df_b_long)

                st.subheader("6. Report anomalie")

                if report.empty:
                    st.success("Nessuna anomalia trovata con i criteri attuali.")
                else:
                    st.warning(f"Sono state trovate {len(report)} possibili anomalie.")
                    st.dataframe(report, use_container_width=True)

                    file_report = crea_excel_report(report, df_b_long)

                    st.download_button(
                        label="Scarica report anomalie in Excel",
                        data=file_report,
                        file_name="report_anomalie.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

else:
    st.warning("Carica entrambi i file Excel per iniziare il confronto.")
