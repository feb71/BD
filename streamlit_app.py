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
                # Se etter 'Fakturanummer' etterfulgt av siffer
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
                    continue

                lines = text.split('\n')
                for line in lines:
                    # Se etter linje som inneholder "Linje", "Artikkel" og "Beløp" (overskrift)
                    if "Linje" in line and "Artikkel" in line and "Beløp" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        tokens = line.split()
                        if len(tokens) < 7:
                            continue

                        # 1) Linjenummer
                        line_num = tokens[0]
                        # Må være tall
                        if not line_num.isdigit():
                            continue

                        # 2) Artikkelnummer (7 siffer)
                        item_number = tokens[1]
                        if not (len(item_number) == 7 and item_number.isdigit()):
                            continue

                        # Siste token = totalpris
                        total_str = tokens[-1].replace('.', '').replace(',', '.')
                        # Nest siste token = enten rabatt eller enhetspris
                        second_last = tokens[-2]
                        discount = None

                        # Testfunksjon: kan streng tolkes som tall?
                        def is_number(s):
                            try:
                                float(s.replace(',', '.').replace('.', ''))
                                return True
                            except ValueError:
                                return False

                        # Hvis nest siste er tall
                        if is_number(second_last):
                            third_last = tokens[-3]
                            if is_number(third_last):
                                # Format: <quantity> <unit> <enhetspris> <rabatt> <total>
                                discount_str = second_last.replace('.', '').replace(',', '.')
                                unit_price_str = third_last.replace('.', '').replace(',', '.')
                                unit = tokens[-4]
                                quantity_str = tokens[-5].replace('.', '').replace(',', '.')

                                discount = float(discount_str)  # Fanger opp rabatten
                                try:
                                    unit_price = float(unit_price_str)
                                    quantity = float(quantity_str)
                                    total_price = float(total_str)
                                except ValueError:
                                    continue

                                desc_tokens = tokens[2:-5]
                            else:
                                # Format: <quantity> <unit> <enhetspris> <total>
                                unit_price_str = second_last.replace('.', '').replace(',', '.')
                                unit = tokens[-3]
                                quantity_str = tokens[-4].replace('.', '').replace(',', '.')

                                try:
                                    unit_price = float(unit_price_str)
                                    quantity = float(quantity_str)
                                    total_price = float(total_str)
                                except ValueError:
                                    continue

                                desc_tokens = tokens[2:-4]
                        else:
                            # Nest siste er ikke tall -> hopp over
                            continue

                        description = " ".join(desc_tokens)
                        unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number

                        data_row = {
                            "UnikID": unique_id,
                            "Varenummer": item_number,
                            "Beskrivelse_Faktura": description,
                            "Antall_Faktura": quantity,
                            "Enhet_Faktura": unit,
                            "Enhetspris_Faktura": unit_price,
                            "Totalt pris": float(total_price),
                            "Type": doc_type
                        }
                        if discount is not None:
                            data_row["Rabatt"] = discount

                        data.append(data_row)

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

def main():
    st.title("Sammenlign Faktura mot Tilbud")
    st.markdown("""<style>.dataframe th {font-weight: bold !important;}</style>""", unsafe_allow_html=True)

    # Opprett tre kolonner
    col1, col2, col3 = st.columns([1, 5, 1])

    with col1:
        st.header("Last opp filer")
        # Tillat opplasting av flere PDF-filer
        invoice_files = st.file_uploader("Last opp fakturaer fra Brødrene Dahl", type="pdf", accept_multiple_files=True)
        offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_files and offer_file:
        all_invoice_data = pd.DataFrame()
        
        # Iterer gjennom alle opplastede PDF-filer
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
                all_invoice_data = pd.concat([all_invoice_data, invoice_data], ignore_index=True)

        # Les tilbudet fra Excel-filen
        with col1:
            st.info("Laster inn tilbud fra Excel-filen...")
        offer_data = pd.read_excel(offer_file)

        # Gi kolonnen navn som i koden
        offer_data.rename(columns={
            'VARENR': 'Varenummer',
            'BESKRIVELSE': 'Beskrivelse_Tilbud',
            'ANTALL': 'Antall_Tilbud',
            'ENHET': 'Enhet_Tilbud',
            'ENHETSPRIS': 'Enhetspris_Tilbud',
            'TOTALPRIS': 'Totalt pris'
        }, inplace=True)

        # Tvinger 'Varenummer' til string i begge DataFrames
        all_invoice_data["Varenummer"] = all_invoice_data["Varenummer"].astype(str)
        offer_data["Varenummer"] = offer_data["Varenummer"].astype(str)

        if not all_invoice_data.empty and not offer_data.empty:
            with col2:
                st.write("Sammenligner data...")

            # Sammenligne faktura mot tilbud
            merged_data = pd.merge(
                offer_data,
                all_invoice_data,
                on="Varenummer",
                how='outer',
                suffixes=('_Tilbud', '_Faktura')
            )

            # Konverter kolonner til numerisk
            merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
            merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
            merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
            merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

            # Finne avvik
            merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
            merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
            merged_data["Prosentvis_økning"] = (
                (merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) 
                / merged_data["Enhetspris_Tilbud"] * 100
            )

            avvik = merged_data[
                (merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))
            ]

            with col2:
                st.subheader("Avvik mellom Faktura og Tilbud")
                st.dataframe(avvik)

            # Artikler som finnes i faktura, men ikke i tilbud
            only_in_invoice = merged_data[merged_data['Enhetspris_Tilbud'].isna()]
            with col2:
                st.subheader("Varenummer som finnes i faktura, men ikke i tilbud")
                st.dataframe(only_in_invoice)

            # Lagre kun artikkeldataene til XLSX
            if not all_invoice_data.empty:
                all_items = all_invoice_data[[
                    "UnikID", 
                    "Varenummer", 
                    "Beskrivelse_Faktura", 
                    "Antall_Faktura", 
                    "Enhetspris_Faktura", 
                    "Totalt pris"
                ]]
            else:
                all_items = pd.DataFrame()

            excel_data_all = convert_df_to_excel(all_items)
            excel_data_avvik = convert_df_to_excel(avvik)
            excel_data_only_invoice = convert_df_to_excel(only_in_invoice)

            with col3:
                st.download_button(
                    label="Last ned avviksrapport (Excel)",
                    data=excel_data_avvik,
                    file_name="avvik_rapport.xlsx"
                )
                
                st.download_button(
                    label="Last ned alle varenummer (Excel)",
                    data=excel_data_all,
                    file_name="faktura_varer.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

                st.download_button(
                    label="Last ned varer kun i faktura (Excel)",
                    data=excel_data_only_invoice,
                    file_name="varer_kun_i_faktura.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        else:
            st.error("Ingen data funnet i de opplastede PDF-filene eller i tilbudet.")
    else:
        st.info("Vennligst last opp både faktura (PDF) og tilbud (Excel).")

if __name__ == "__main__":
    main()
