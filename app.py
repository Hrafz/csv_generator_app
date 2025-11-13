import streamlit as st
import pandas as pd
import os
import csv
from io import BytesIO
import zipfile

def generate_csv_files(excel_file):
    """Generate CSV files from Excel sheets and return them as a dictionary"""
    # Define the list of sheet names you want to process
    sheet_names = ["BU25 - ORDER MASS PROD ", "BU25 - SPECIFIC ORDERS", "BU POS", "BU POS SPE"]

    csv_files = {}
    results = []

    # Loop through the specified sheets
    for sheet_name in sheet_names:
        try:
            # Read the sheet into a DataFrame
            df = pd.read_excel(excel_file, sheet_name=sheet_name)

            if df.empty:
                results.append(f"⚠️ The sheet '{sheet_name}' appears to be empty.")
                continue

            # Determine columns to be emptied based on the header row
            if sheet_name in ["BU POS", "BU POS SPE"]:
                empty_columns = [col for col in df.columns if str(col).startswith('Unnamed')]
                df = df.drop(columns=empty_columns)

            # Convert DataFrame to CSV string
            csv_buffer = BytesIO()
            df.to_csv(csv_buffer, index=False, quoting=csv.QUOTE_ALL, encoding='utf-8')
            csv_files[sheet_name + ".csv"] = csv_buffer.getvalue()

            results.append(f"✅ CSV file for sheet '{sheet_name}' has been generated successfully!")

        except Exception as e:
            results.append(f"❌ An error occurred while processing sheet '{sheet_name}': {e}")

    return csv_files, results

def create_zip(csv_files):
    """Create a zip file containing all CSV files"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, content in csv_files.items():
            zip_file.writestr(filename, content)
    return zip_buffer.getvalue()

# Streamlit App
st.set_page_config(
    page_title="Excel to CSV Converter",
    page_icon="📊",
    layout="centered"
)

st.title("📊 Excel to CSV Converter")
st.write("Upload an Excel file to convert specific sheets to CSV format.")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xlsm', 'xls'])

if uploaded_file is not None:
    st.info(f"File uploaded: {uploaded_file.name}")

    # Add a convert button
    if st.button("Convert to CSV", type="primary"):
        with st.spinner("Processing Excel file..."):
            try:
                # Generate CSV files
                csv_files, results = generate_csv_files(uploaded_file)

                # Display results
                st.subheader("Processing Results:")
                for result in results:
                    st.write(result)

                if csv_files:
                    st.success(f"Successfully generated {len(csv_files)} CSV file(s)!")

                    # Create two columns for download options
                    col1, col2 = st.columns(2)

                    with col1:
                        # Download all as zip
                        zip_data = create_zip(csv_files)
                        st.download_button(
                            label="📦 Download All as ZIP",
                            data=zip_data,
                            file_name="csv_files.zip",
                            mime="application/zip"
                        )

                    # Individual file downloads
                    st.subheader("Download Individual Files:")
                    for filename, content in csv_files.items():
                        st.download_button(
                            label=f"⬇️ {filename}",
                            data=content,
                            file_name=filename,
                            mime="text/csv"
                        )
                else:
                    st.warning("No CSV files were generated. Please check the Excel file.")

            except Exception as e:
                st.error(f"An error occurred: {e}")
else:
    st.info("👆 Please upload an Excel file to get started.")

    # Show expected sheets
    with st.expander("ℹ️ Expected Sheet Names"):
        st.write("The app will process the following sheets:")
        st.write("- BU25 - ORDER MASS PROD")
        st.write("- BU25 - SPECIFIC ORDERS")
        st.write("- BU POS")
        st.write("- BU POS SPE")
