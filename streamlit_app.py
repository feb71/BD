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

# Korrekt tilpasset funksjon for å lese PDF-faktura fra Brødrene Dahl AS
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
                    if re.search(r"Linje\s+Artikkel.*BeløpBeskrivelse", line):
                        start_reading = True
                        continue

                    if start_reading:
                        columns = line.split()
                        if len(columns) >= 7:
                            line_num = columns[0]
                            item_number = columns[1]

                            total_price = columns[-1].replace('.', '').replace(',', '.')
                            unit_price = columns[-2].replace('.', '').replace(',', '.')
                            unit = columns[-3]
                            quantity = columns[-4].replace('.', '').replace(',', '.')

                            description = " ".join(columns[2:-4])

                            unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number

                            data.append({
                                "UnikID": unique_id,
                                "Varenummer": item_number,
                                "Beskrivelse_Faktura": description,
                                "Antall_Faktura": float(quantity),
                                "Enhet_Faktura": unit,
                                "Enhetspris_Faktura": float(unit_price),
                                "Totalt pris": float(total_price),
                                "Type": doc_type
                            })

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

        offer_data = pd.read_excel(offer_file)
        offer_data.rename(columns={
            'VARENR': 'Varenummer',
            'BESKRIVELSE': 'Beskrivelse_Tilbud',
            'ANTALL': 'Antall_Tilbud',
            'ENHET': 'Enhet_Tilbud',
            'ENHETSPRIS': 'Enhetspris_Tilbud',
            'TOTALPRIS': 'Totalt pris'
        }, inplace=True)

        if not all_invoice_data.empty and not offer_data.empty:
            merged_data = pd.merge(offer_data, all_invoice_data, on="Varenummer", how='outer', suffixes=('_Tilbud', '_Faktura'))

            st.dataframe(merged_data)

            excel_data = convert_df_to_excel(merged_data)

            st.download_button(
                label="Last ned sammenlignet data som Excel",
                data=excel_data,
                file_name="sammenlignet_data.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.error("Ingen data funnet i de opplastede filene.")

if __name__ == "__main__":
    main()
