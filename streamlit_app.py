import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Streamlit App", layout="wide", initial_sidebar_state="expanded")

# Funksjon for å lese fakturanummer fra PDF
def get_invoice_number(file):
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                match = re.search(r"Fakturanummer\s*[:\-]?\s*(\d+)", text, re.IGNORECASE)
                if match:
                    return match.group(1)
        return None
    except Exception as e:
        st.error(f"Kunne ikke lese fakturanummer fra PDF: {e}")
        return None

# Funksjon for å lese PDF-filen og hente ut relevante data
def extract_data_from_pdf(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    st.error(f"Ingen tekst funnet på side {page.page_number} i PDF-filen.")
                    continue

                lines = text.split('\n')
                for line in lines:
                    # Start reading when we encounter the header line
                    if doc_type == "Faktura" and "Artikkel" in line and "Beløp" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        if not line.strip():  # skip empty lines
                            continue
                        tokens = line.split()
                        if not tokens:
                            continue
                        if not tokens[0].isdigit():
                            # Skip lines that do not start with a line number (e.g. summaries or footer text)
                            continue
                        if len(tokens) < 2:
                            continue
                        item_token = tokens[1]
                        if not (len(item_token) == 7 and item_token.isdigit()):
                            # Skip line if second token is not a 7-digit item number
                            continue

                        # Rens tokens (fjern '%' tegn som eget token eller suffiks)
                        tokens_clean = []
                        for t in tokens:
                            if t == '%':
                                continue
                            if t.endswith('%'):
                                t = t[:-1]
                            tokens_clean.append(t)
                        tokens = tokens_clean
                        if len(tokens) < 6:
                            # Invoice line should have at least 6 tokens (linje, artikkel, etc.)
                            continue

                        # Sjekk de siste tokens for å identifisere pris og beløp (og rabatt om den finnes)
                        def is_numeric_token(tok):
                            try:
                                float(tok.replace('.', '').replace(',', '.'))
                                return True
                            except:
                                return False

                        # Sjekk at de to siste tokens er numeriske (salgspris og beløp)
                        if not (is_numeric_token(tokens[-1]) and is_numeric_token(tokens[-2])):
                            continue

                        # Hvis tredje siste token også er numerisk, antar vi at rabatt% er til stede
                        if is_numeric_token(tokens[-3]):
                            # Linjen inkluderer rabatt%
                            total_str = tokens[-1]
                            discount_str = tokens[-2]  # rabattprosent (ikke brukt videre)
                            price_str = tokens[-3]
                            unit_str = tokens[-4]
                            quantity_str = tokens[-5]
                            desc_tokens = tokens[2:-5]
                        else:
                            # Linjen har ingen rabatt-kolonne
                            total_str = tokens[-1]
                            price_str = tokens[-2]
                            unit_str = tokens[-3]
                            quantity_str = tokens[-4]
                            discount_str = None
                            desc_tokens = tokens[2:-4]

                        description = " ".join(desc_tokens)

                        # Konverter tallverdier fra tekst til float
                        def parse_number(num_str):
                            try:
                                return float(num_str.replace('.', '').replace(',', '.'))
                            except:
                                return None

                        quantity = parse_number(quantity_str)
                        unit_price = parse_number(price_str)
                        total_price = parse_number(total_str)
                        if quantity is None or unit_price is None or total_price is None:
                            # Hopp over linjen hvis vi ikke klarer å lese tallverdiene
                            continue

                        item_number = str(item_token)
                        unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                        data.append({
                            "UnikID": unique_id,
                            "Varenummer": item_number,
                            "Beskrivelse_Faktura": description,
                            "Antall_Faktura": quantity,
                            "Enhet_Faktura": unit_str,
                            "Enhetspris_Faktura": unit_price,
                            "Totalt pris": total_price,
                            "Type": doc_type
                        })
            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")

            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å konvertere DataFrame til en Excel-fil
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")
    st.markdown("""<style>.dataframe th {font-weight: bold !important;}</style>""", unsafe_allow_html=True)

    # Opprett tre kolonner for layout
    col1, col2, col3 = st.columns([1, 5, 1])

    with col1:
        st.header("Last opp filer")
        # Tillat opplasting av flere PDF-filer (fakturaer)
        invoice_files = st.file_uploader("Last opp fakturaer fra Brødrene Dahl", type="pdf", accept_multiple_files=True)
        # Last opp én Excel-fil (tilbud)
        offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_files and offer_file:
        all_invoice_data = pd.DataFrame()
        # Behandle hver opplastet faktura-PDF
        for invoice_file in invoice_files:
            with col1:
                st.info(f"Henter fakturanummer fra {invoice_file.name}...")
            invoice_number = get_invoice_number(invoice_file)

            if invoice_number:
                with col1:
                    st.success(f"Fakturanummer funnet: {invoice_number}")
                # Ekstraher data fra PDF-filen
                with col1:
                    st.info(f"Laster inn faktura fra {invoice_file.name}...")
                invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)
                # Legg til data fra denne fakturaen i samlet DataFrame
                all_invoice_data = pd.concat([all_invoice_data, invoice_data], ignore_index=True)

        # Les tilbudsdata fra Excel-filen
        with col1:
            st.info("Laster inn tilbud fra Excel-filen...")
        offer_data = pd.read_excel(offer_file)
        offer_data.rename(columns={
            'VARENR': 'Varenummer',
            'BESKRIVELSE': 'Beskrivelse_Tilbud',
            'ANTALL': 'Antall_Tilbud',
            'ENHET': 'Enhet_Tilbud',
            'ENHETSPRIS': 'Enhetspris_Tilbud',
            'TOTALPRIS': 'Totalt pris'
        }, inplace=True)

        # Sørg for at 'Varenummer' er tekst i begge DataFrames (for korrekt merge)
        if not all_invoice_data.empty:
            all_invoice_data['Varenummer'] = all_invoice_data['Varenummer'].astype(str)
        if not offer_data.empty:
            offer_data['Varenummer'] = offer_data['Varenummer'].astype(str)

        if not offer_data.empty:
            # Slå sammen faktura- og tilbudsdata på Varenummer
            with col2:
                st.write("Sammenligner data...")
            merged_data = pd.merge(offer_data, all_invoice_data, on="Varenummer", how='outer', suffixes=('_Tilbud', '_Faktura'))

            # Konverter kvantum og prisfelt til numerisk for beregning
            merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
            merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
            merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
            merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

            # Beregn avvik og prosentvis økning
            merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
            merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
            merged_data["Prosentvis_økning"] = ((merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) / merged_data["Enhetspris_Tilbud"]) * 100

            # Filtrer ut avvik (der det faktisk er forskjeller)
            avvik = merged_data[
                (merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))
            ]

            with col2:
                st.subheader("Avvik mellom Faktura og Tilbud")
                st.dataframe(avvik)

            # Finn varer som kun finnes på faktura (ikke i tilbud)
            only_in_invoice = merged_data[merged_data['Enhetspris_Tilbud'].isna()]
            with col2:
                st.subheader("Varenummer som finnes i faktura, men ikke i tilbud")
                st.dataframe(only_in_invoice)

            # Gjør klar alle fakturalinjer for nedlasting (Excel)
            all_items = all_invoice_data[["UnikID", "Varenummer", "Beskrivelse_Faktura", "Antall_Faktura", "Enhetspris_Faktura", "Totalt pris"]]
            excel_data = convert_df_to_excel(all_items)

            with col3:
                st.download_button(
                    label="Last ned avviksrapport som Excel",
                    data=convert_df_to_excel(avvik),
                    file_name="avvik_rapport.xlsx"
                )
                st.download_button(
                    label="Last ned alle varenummer som Excel",
                    data=excel_data,
                    file_name="faktura_varer.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                # Lag Excel-fil med varer som kun finnes i fakturaen (ikke i tilbudet)
                only_in_invoice_data = convert_df_to_excel(only_in_invoice)
                st.download_button(
                    label="Last ned varenummer som ikke eksiterer i tilbudet",
                    data=only_in_invoice_data,
                    file_name="varer_kun_i_faktura.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        else:
            st.error("Kunne ikke lese tilbudsdata fra Excel-filen.")

if __name__ == "__main__":
    main()
