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

# Tilpasset funksjon for å lese PDF-faktura fra Brødrene Dahl AS (robust versjon)
def extract_data_from_pdf(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    continue

                lines = text.split('\n')
                for line in lines:
                    # Sjekk om vi er på overskriftslinjen
                    if "Linje" in line and "Artikkel" in line and "Beløp" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        columns = line.split()

                        # For å unngå at vi prøver å parse tekstlinjer med for få kolonner
                        if len(columns) < 7:
                            continue

                        # Pass på at første kolonne er siffer (linjenummer)
                        if not columns[0].isdigit():
                            continue

                        # Pakk konvertering til float i try/except
                        try:
                            line_num = columns[0]
                            item_number = columns[1]

                            total_price = columns[-1].replace('.', '').replace(',', '.')
                            unit_price = columns[-2].replace('.', '').replace(',', '.')
                            unit = columns[-3]
                            quantity = columns[-4].replace('.', '').replace(',', '.')

                            description = " ".join(columns[2:-4])

                            # Konverter til float
                            total_price = float(total_price)
                            unit_price = float(unit_price)
                            quantity = float(quantity)

                            unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number

                            data.append({
                                "UnikID": unique_id,
                                "Varenummer": item_number,
                                "Beskrivelse_Faktura": description,
                                "Antall_Faktura": quantity,
                                "Enhet_Faktura": unit,
                                "Enhetspris_Faktura": unit_price,
                                "Totalt pris": total_price,
                                "Type": doc_type
                            })
                        except ValueError:
                            # Hopper over linjen om konvertering feiler
                            continue

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
    st.title("Les og sammenlign faktura med tilbud fra Brødrene Dahl")

    invoice_files = st.file_uploader("Last opp fakturaer fra Brødrene Dahl", type="pdf", accept_multiple_files=True)
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_files and offer_file:
        all_invoice_data = pd.DataFrame()

        for invoice_file in invoice_files:
            invoice_number = get_invoice_number(invoice_file)

            if invoice_number:
                invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)
                all_invoice_data = pd.concat([all_invoice_data, invoice_data], ignore_index=True)

        # Les tilbudet fra Excel-filen
        offer_data = pd.read_excel(offer_file)
        offer_data.rename(columns={
            'VARENR': 'Varenummer',
            'BESKRIVELSE': 'Beskrivelse_Tilbud',
            'ANTALL': 'Antall_Tilbud',
            'ENHET': 'Enhet_Tilbud',
            'ENHETSPRIS': 'Enhetspris_Tilbud',
            'TOTALPRIS': 'Totalt pris'
        }, inplace=True)

        # Sjekk at både PDF-data og tilbud har innhold
        if not all_invoice_data.empty and not offer_data.empty:
            # Slå sammen basert på 'Varenummer'
            merged_data = pd.merge(
                offer_data,
                all_invoice_data,
                on="Varenummer",
                how='outer',
                suffixes=('_Tilbud', '_Faktura')
            )

            st.dataframe(merged_data)

            # Gjør det mulig å laste ned resultatet
            excel_data = convert_df_to_excel(merged_data)
            st.download_button(
                label="Last ned sammenlignet data som Excel",
                data=excel_data,
                file_name="sammenlignet_data.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.error("Ingen data funnet i de opplastede filene.")
    else:
        st.info("Vennligst last opp både faktura(er) (PDF) og tilbud (Excel).")

if __name__ == "__main__":
    main()
