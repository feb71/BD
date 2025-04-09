import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Streamlit App", layout="wide", initial_sidebar_state="expanded")

# Funksjon for å lese fakturanummer fra PDF
def get_invoice_number(file):
    """
    Søker gjennom PDF-en og finner fakturanummer, 
    basert på 'Fakturanummer' + siffer.
    """
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

# Denne funksjonen er oppdatert for å takle den nye PDF-strukturen
def extract_data_from_pdf(file, doc_type, invoice_number=None):
    """
    Les PDF-faktura og ekstraher varelinjer.
    - Oppdager om linjen har en rabattkolonne (f.eks. '45,00') 
      eller ikke, ved å sjekke de siste tokenene.
    - Linjer som ikke passer formatet, hoppes over.
    - 'UnikID' = <invoice_number>_<varenummer>, 
      for å entydig identifisere linjene.
    """
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False  # Blir 'True' når vi finner overskriften

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    continue

                lines = text.split('\n')
                for line in lines:
                    # Oppdag overskriften som inneholder "Linje", "Artikkel" og "Beløp"
                    if "Linje" in line and "Artikkel" in line and "Beløp" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        # Del opp linjen i tokens (kolonner)
                        tokens = line.split()
                        if len(tokens) < 7:
                            continue

                        # 1) Linjenummer
                        line_num = tokens[0]
                        # Må være tall
                        if not line_num.isdigit():
                            continue

                        # 2) Artikkelnummer
                        item_number = tokens[1]
                        # Sjekk at det er 7 siffer (Brødrene Dahl-type varenr)
                        if not (len(item_number) == 7 and item_number.isdigit()):
                            continue

                        # Funksjon: teste om streng kan tolkes som et tall
                        def is_number(s):
                            try:
                                float(s.replace('.', '').replace(',', '.'))
                                return True
                            except ValueError:
                                return False

                        # Siste token = totalpris
                        total_str = tokens[-1].replace('.', '').replace(',', '.')
                        second_last = tokens[-2]  # nest siste token
                        discount = None  # default ingen rabatt

                        # Avhengig av om nest siste er tall → enten rabatt eller enhetspris
                        if is_number(second_last):
                            # Tredje siste kan være enhetspris eller rabatt
                            third_last = tokens[-3]
                            if is_number(third_last):
                                # Da har vi: 
                                # discount = second_last
                                # enhetspris = third_last
                                # unit = tokens[-4]
                                # quantity = tokens[-5]
                                discount_str = second_last.replace('.', '').replace(',', '.')
                                unit_price_str = third_last.replace('.', '').replace(',', '.')
                                unit = tokens[-4]
                                quantity_str = tokens[-5].replace('.', '').replace(',', '.')
                                desc_tokens = tokens[2:-5]  # alt mellom item_number og quantity

                                try:
                                    discount = float(discount_str)
                                    unit_price = float(unit_price_str)
                                    quantity = float(quantity_str)
                                    total_price = float(total_str)
                                except ValueError:
                                    # Hvis konvertering feilet, hopp over linjen
                                    continue
                            else:
                                # Ingen rabatt. second_last = enhetspris
                                # third_last = unit, 
                                # fourth_last = quantity
                                unit_price_str = second_last.replace('.', '').replace(',', '.')
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
                            # Nest siste er ikke tall, 
                            # da er formatet annerledes enn forventet
                            continue

                        # Beskrivelse: alt mellom tokens[2] og tokens[-4], 
                        # ev. -5 hvis vi har rabatt
                        description = " ".join(desc_tokens)
                        unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number

                        # Bygg rad
                        row = {
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
                            row["Rabatt"] = discount

                        data.append(row)

            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

def convert_df_to_excel(df):
    """
    Konverterer en Pandas DataFrame til Excel-format (i minnet) 
    og returnerer en BytesIO.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def main():
    st.title("Sammenlign Faktura mot Tilbud")
    st.markdown("""<style>.dataframe th {font-weight: bold !important;}</style>""", unsafe_allow_html=True)

    # Opprett to opplastingsfelter: PDF-faktura(er) + Excel-tilbud
    invoice_files = st.file_uploader("Last opp fakturaer fra Brødrene Dahl", type="pdf", accept_multiple_files=True)
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_files and offer_file:
        # Samler data for alle PDF-filer
        all_invoice_data = pd.DataFrame()

        # Gå gjennom hver PDF-faktura
        for invoice_file in invoice_files:
            invoice_number = get_invoice_number(invoice_file)
            if invoice_number:
                invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)
                all_invoice_data = pd.concat([all_invoice_data, invoice_data], ignore_index=True)

        # Les tilbudsdata fra Excel
        offer_data = pd.read_excel(offer_file)
        offer_data.rename(columns={
            'VARENR': 'Varenummer',
            'BESKRIVELSE': 'Beskrivelse_Tilbud',
            'ANTALL': 'Antall_Tilbud',
            'ENHET': 'Enhet_Tilbud',
            'ENHETSPRIS': 'Enhetspris_Tilbud',
            'TOTALPRIS': 'Totalt pris'
        }, inplace=True)

        # Sørg for at vi kan flette dataene
        all_invoice_data["Varenummer"] = all_invoice_data["Varenummer"].astype(str)
        offer_data["Varenummer"] = offer_data["Varenummer"].astype(str)

        if not all_invoice_data.empty and not offer_data.empty:
            # Slå sammen på varenummer
            merged_data = pd.merge(
                offer_data,
                all_invoice_data,
                on="Varenummer",
                how='outer',
                suffixes=('_Tilbud', '_Faktura')
            )

            st.subheader("Sammenslått tabell")
            st.dataframe(merged_data)

            # Tilby nedlasting av resultatet i Excel
            excel_data = convert_df_to_excel(merged_data)
            st.download_button(
                label="Last ned sammenlignet data som Excel",
                data=excel_data,
                file_name="sammenlignet_data.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.error("Ingen data funnet i de opplastede PDF-filene eller i tilbudsfilen.")
    else:
        st.info("Last opp både faktura (PDF) og tilbud (Excel) for å sammenligne.")

if __name__ == "__main__":
    main()
