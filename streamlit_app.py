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
                match = re.search(r"Fakturanummer\\s*[:\\-]?\\s*(\\d+)", text, re.IGNORECASE)
                if match:
                    return match.group(1)
        return None
    except Exception as e:
        st.error(f"Kunne ikke lese fakturanummer fra PDF: {e}")
        return None

# Ny, robust funksjon som håndterer valgfri rabatt-kolonne
# Varenummer + Fakturanummer = UnikID

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
                    # Oppdag overskriften
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

                        # 2) Artikkelnummer
                        item_number = tokens[1]
                        # Sjekk at artikkelnummeret er 7 siffer
                        if not (len(item_number) == 7 and item_number.isdigit()):
                            continue

                        # Siste token er totalpris
                        total_str = tokens[-1].replace('.', '').replace(',', '.')

                        # Vi sjekker om nest siste token er rabatt eller enhetspris
                        second_last = tokens[-2]

                        def is_number(s):
                            try:
                                float(s.replace(',', '.').replace('.', ''))
                                return True
                            except:
                                return False

                        discount = None

                        if is_number(second_last):
                            # Tredje siste kan være enhetspris eller rabatt
                            third_last = tokens[-3]
                            if is_number(third_last):
                                # Da har vi rabatten = second_last,
                                # enhetspris = third_last,
                                # fjerde siste = unit,
                                # femte siste = quantity
                                discount_str = second_last.replace('.', '').replace(',', '.')
                                unit_price_str = third_last.replace('.', '').replace(',', '.')
                                unit = tokens[-4]
                                quantity_str = tokens[-5].replace('.', '').replace(',', '.')
                                desc_tokens = tokens[2:-5]

                                discount = float(discount_str)

                                try:
                                    unit_price = float(unit_price_str)
                                    quantity = float(quantity_str)
                                    total_price = float(total_str)
                                except ValueError:
                                    continue
                            else:
                                # Ingen rabatt, second_last er enhetspris,
                                unit_price_str = second_last
                                unit = tokens[-3]
                                quantity_str = tokens[-4].replace('.', '').replace(',', '.')
                                desc_tokens = tokens[2:-4]

                                try:
                                    unit_price = float(unit_price_str)
                                    quantity = float(quantity_str)
                                    total_price = float(total_str)
                                except ValueError:
                                    continue
                        else:
                            # Nest siste er ikke tall => format avviker
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


def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()


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

        # For å unngå merge-feil
        all_invoice_data["Varenummer"] = all_invoice_data["Varenummer"].astype(str)
        offer_data["Varenummer"] = offer_data["Varenummer"].astype(str)

        if not all_invoice_data.empty and not offer_data.empty:
            merged_data = pd.merge(
                offer_data,
                all_invoice_data,
                on="Varenummer",
                how='outer',
                suffixes=('_Tilbud', '_Faktura')
            )

            st.subheader("Sammenslått tabell")
            st.dataframe(merged_data)

            # Opprett tabell 1: Varenummer som finnes i tilbudet
            table_in_offer = merged_data[~merged_data['Beskrivelse_Tilbud'].isna()]

            # Opprett tabell 2: Varenummer som IKKE finnes i tilbudet
            table_not_in_offer = merged_data[merged_data['Beskrivelse_Tilbud'].isna()]

            st.subheader("Varer som finnes i tilbudet")
            st.dataframe(table_in_offer)

            st.subheader("Varer som IKKE finnes i tilbudet")
            st.dataframe(table_not_in_offer)

            # Gjør dem nedlastbare
            excel_data_1 = convert_df_to_excel(table_in_offer)
            excel_data_2 = convert_df_to_excel(table_not_in_offer)

            st.download_button(
                label="Last ned (Excel) - Varer i tilbudet",
                data=excel_data_1,
                file_name="varer_i_tilbudet.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

            st.download_button(
                label="Last ned (Excel) - Varer IKKE i tilbudet",
                data=excel_data_2,
                file_name="varer_ikke_i_tilbudet.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.error("Ingen data funnet i de opplastede filene.")
    else:
        st.info("Vennligst last opp både faktura (PDF) og tilbud (Excel).")

if __name__ == "__main__":
    main()
