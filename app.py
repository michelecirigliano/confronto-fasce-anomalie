import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Confronto fasce e anomalie",
    layout="wide"
)

st.title("Confronto fasce e anomalie")
st.write(
    "Carica due file Excel per confrontare date, nominativi e fasce orarie. "
    "Questa è una prima versione tecnica dell'app."
)

file_a = st.file_uploader("Carica il primo file Excel", type=["xlsx"], key="file_a")
file_b = st.file_uploader("Carica il secondo file Excel", type=["xlsx"], key="file_b")

def carica_excel(file):
    try:
        return pd.read_excel(file)
    except Exception as e:
        st.error(f"Errore nella lettura del file: {e}")
        return None

def crea_report_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Report anomalie")
    output.seek(0)
    return output

if file_a is not None and file_b is not None:
    df_a = carica_excel(file_a)
    df_b = carica_excel(file_b)

    if df_a is not None and df_b is not None:
        st.subheader("Anteprima file 1")
        st.dataframe(df_a.head(20), use_container_width=True)

        st.subheader("Anteprima file 2")
        st.dataframe(df_b.head(20), use_container_width=True)

        st.subheader("Colonne rilevate")

        col1, col2 = st.columns(2)

        with col1:
            st.write("Colonne file 1")
            st.write(list(df_a.columns))

        with col2:
            st.write("Colonne file 2")
            st.write(list(df_b.columns))

        st.info(
            "Questa prima versione serve a verificare che l'app legga correttamente "
            "i due file Excel. Nel passaggio successivo aggiungeremo la scelta delle "
            "colonne e il motore di confronto delle sovrapposizioni."
        )

        report_base = pd.DataFrame({
            "Esito": ["File caricati correttamente"],
            "Righe file 1": [len(df_a)],
            "Righe file 2": [len(df_b)]
        })

        report_excel = crea_report_excel(report_base)

        st.download_button(
            label="Scarica report di prova",
            data=report_excel,
            file_name="report_prova.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("Carica entrambi i file Excel per iniziare.")
