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

# Oppdatert funksjon for å lese PDF-faktura fra Brødrene Dahl
# Håndterer valgfri rabattkolonne og lager UnikID av fakturanummer og varenummer
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
                    # Oppdag overskriften (Linje, Artikkel, Beløp)
                    if "Linje" in line and "Artikkel" in line and "Beløp" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        tokens = line.split()
                        if len(tokens) < 7:
                            continue

                        # 1) Linjenummer
                        line_num = tokens[0]
                        if not line_num.isdigit():
                            continue

                        # 2) Artikkelnummer (7 siffer)
                        item_number = tokens[1]
                        if not (len(item_number) == 7 and item_number.isdigit()):
                            continue

                        # Funksjon for å teste om streng kan være et tall
                        def is_number(s):
                            try:
                                float(s.replace(',', '.').replace('.', ''))
                                return True
                            except ValueError:
                                return False

                        # Siste token i linjen -> totalpris
                        total_str = tokens[-1].replace('.', '').replace(',', '.')
                        # Nest siste token -> enten rabatt eller enhetspris
                        second_last = tokens[-2]
                        discount = None

                        # Sjekk om nest siste er tall
                        if is_number(second_last):
                            # Tredje siste kan da være enhetspris eller rabatt
                            third_last = tokens[-3]
                            if is_number(third_last):
                                # Format: ... <quantity> <unit> <enhetspris> <rabatt> <total>
                                discount_str = second_last.replace('.', '').replace(',', '.')
                                unit_price_str = third_last.replace('.', '').replace(',', '.')
                                unit = tokens[-4]
                                quantity_str = tokens[-5].replace('.', '').replace(',', '.')

                                discount = float(discount_str)  # vi tar den med
                                try:
                                    unit_price = float(unit_price_str)
                                    quantity = float(quantity_str)
                                    total_price = float(total_str)
                                except ValueError:
                                    continue

                                desc_tokens = tokens[2:-5]
                            else:
                                # Format: ... <quantity> <unit> <enhetspris> <total>
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
                            "Totalt pris": total_price,
                            "Type": doc_type
                        }
                        if discount is not None:
                            data_row["Rabatt"] = discount

                        data.append(data_row)

            return pd.DataFrame(data)

    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Konverterer DataFrame til Excel
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def main():
    st.title("Sammenlign Faktura mot Tilbud")

    col1, col2 = st.columns([1,3])

    with col1:
        # Last opp PDF-filer (faktura) og Excel-fil (tilbud)
        invoice_files = st.file_uploader("Last opp fakturaer (PDF)", type="pdf", accept_multiple_files=True)
        offer_file = st.file_uploader("Last opp tilbud (Excel)", type="xlsx")

    if invoice_files and offer_file:
        all_invoice_data = pd.DataFrame()

        # Les PDF-filer
        for invoice_file in invoice_files:
            invoice_number = get_invoice_number(invoice_file)
            if invoice_number:
                with col1:
                    st.write(f"Fakturanummer funnet: {invoice_number}")
                invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)
                all_invoice_data = pd.concat([all_invoice_data, invoice_data], ignore_index=True)

        # Les tilbud fra Excel
        offer_data = pd.read_excel(offer_file)
        offer_data.rename(columns={
            'VARENR': 'Varenummer',
            'BESKRIVELSE': 'Beskrivelse_Tilbud',
            'ANTALL': 'Antall_Tilbud',
            'ENHET': 'Enhet_Tilbud',
            'ENHETSPRIS': 'Enhetspris_Tilbud',
            'TOTALPRIS': 'Totalt pris'
        }, inplace=True)

        # Tvinger Varenummer til samme datatype (string)
        all_invoice_data["Varenummer"] = all_invoice_data["Varenummer"].astype(str)
        offer_data["Varenummer"] = offer_data["Varenummer"].astype(str)

        if not all_invoice_data.empty and not offer_data.empty:
            # Slå sammen data på Varenummer
            merged_data = pd.merge(
                offer_data,
                all_invoice_data,
                on="Varenummer",
                how='outer',
                suffixes=('_Tilbud', '_Faktura')
            )

            # Gjør kolonner numeriske for sammenligning
            merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
            merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
            merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
            merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

            # Beregn avvik
            merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
            merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
            merged_data["Prosentvis_økning"] = (
                (merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) 
                / merged_data["Enhetspris_Tilbud"] * 100
            )

            # Tabell: Varer som finnes i tilbudet (der Enhetspris_Tilbud finnes)
            table_in_offer = merged_data[ merged_data["Enhetspris_Tilbud"].notna() ]

            # Tabell: Varer som IKKE finnes i tilbudet (der Enhetspris_Tilbud er NaN)
            table_not_in_offer = merged_data[ merged_data["Enhetspris_Tilbud"].isna() ]

            with col2:
                st.subheader("Alle sammenlignede data:")
                st.dataframe(merged_data)

                st.subheader("Varer som finnes i tilbudet:")
                st.dataframe(table_in_offer)

                st.subheader("Varer som IKKE finnes i tilbudet:")
                st.dataframe(table_not_in_offer)

            # Gjør dem nedlastbare
            excel_data_merged = convert_df_to_excel(merged_data)
            excel_data_in = convert_df_to_excel(table_in_offer)
            excel_data_out = convert_df_to_excel(table_not_in_offer)

            with col1:
                st.download_button(
                    label="Last ned all data (Excel)",
                    data=excel_data_merged,
                    file_name="alle_varer.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                st.download_button(
                    label="Last ned varer i tilbudet (Excel)",
                    data=excel_data_in,
                    file_name="varer_i_tilbudet.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                st.download_button(
                    label="Last ned varer IKKE i tilbudet (Excel)",
                    data=excel_data_out,
                    file_name="varer_ikke_i_tilbudet.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        else:
            st.error("Ingen data funnet i enten PDF- eller Excel-filen. Sjekk at de inneholder riktige kolonner/verdier.")
    else:
        st.info("Vennligst last opp både faktura (PDF) og tilbud (Excel) for å sammenligne.")

if __name__ == "__main__":
    main()
