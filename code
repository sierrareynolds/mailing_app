import streamlit as st
import pandas as pd

st.set_page_config(page_title="Mopit Mailing List Generator", layout="centered")
st.title("📬 Mopit Mailing List Generator")
st.markdown("Upload address files and generate a mailing list for companies that do not own a Mopit machine.")

# Uploaders and input
file_main = st.file_uploader("Upload Main Company Address Excel File", type=[".xls", ".xlsx"])
file_ship = st.file_uploader("Upload Machine Owner Address Excel File", type=[".xls", ".xlsx"])
company_name = st.text_input("Enter Company Name (used in file name)")

if st.button("Generate Mailing List"):
    if not (file_main and file_ship and company_name):
        st.warning("Please upload both files and enter a company name.")
    else:
        try:
            df = pd.read_excel(file_main)
            df_ship = pd.read_excel(file_ship)

            # Filter US addresses and clean ZIP codes
            df = df[df['country_code'] == 'US'].copy()
            df['zip_code'] = df['zip_code'].astype(str).str.extract(r'(\d+)')[0].fillna('00000').str.zfill(5)

            # Construct full address
            df['full_address'] = (
                df['address'].str.upper() + ', ' +
                df['city'].str.upper() + ', ' +
                df['state'].str.upper() + ' ' +
                df['zip_code'].str.upper()
            )

            # Process shipping addresses
            cols_ship = ['Ship To Street1', 'Ship To City', 'Ship To State', 'Ship To Zip']
            df_ship[cols_ship] = df_ship[cols_ship].astype(str)
            df_ship['full_ship_to_address'] = (
                df_ship['Ship To Street1'].str.upper() + ', ' +
                df_ship['Ship To City'].str.upper() + ', ' +
                df_ship['Ship To State'].str.upper() + ' ' +
                df_ship['Ship To Zip'].str.upper()
            )

            # Find non-purchasing addresses
            df_not_purchased = df[~df['full_address'].isin(df_ship['full_ship_to_address'])]

            # Export to Excel in memory
            output_filename = f"{company_name}_mailing_list.xlsx"
            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                df_not_purchased.to_excel(writer, index=False)

            with open(output_filename, 'rb') as f:
                st.success("Mailing list created!")
                st.download_button("📥 Download Mailing List", f, file_name=output_filename)

        except Exception as e:
            st.error(f"An error occurred: {e}")
