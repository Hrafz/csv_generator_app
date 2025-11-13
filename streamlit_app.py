import streamlit as st
import pandas as pd
import csv
from io import BytesIO
import zipfile
import base64

st.set_page_config(
    page_title="CSV Generator from Excel",
    page_icon="📊",
    layout="wide"
)

# Custom CSS for button color and compact layout
st.markdown("""
    <style>
    .stButton > button[kind="primary"] {
        background-color: #0047AB;
        border-color: #0047AB;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #003580;
        border-color: #003580;
    }
    .stCheckbox {
        margin-bottom: 0rem !important;
    }
    div[data-testid="column"] {
        padding: 0.25rem 0.5rem;
    }
    h1 {
        margin-top: 0rem;
        margin-bottom: 1rem;
    }
    h3 {
        margin-top: 0.5rem;
        margin-bottom: 0.5rem;
    }
    </style>
    """, unsafe_allow_html=True)

# Initialize session state
if 'available_sheets' not in st.session_state:
    st.session_state.available_sheets = []
if 'selected_sheets' not in st.session_state:
    st.session_state.selected_sheets = {}
if 'bupos_sheets' not in st.session_state:
    st.session_state.bupos_sheets = []
if 'log_messages' not in st.session_state:
    st.session_state.log_messages = []

def add_log(message):
    """Add message to log"""
    st.session_state.log_messages.append(message)

def generate_csv_files(excel_file):
    """Generate CSV files from selected sheets"""
    csv_files = {}
    st.session_state.log_messages = []

    selected_sheet_names = [sheet for sheet, selected in st.session_state.selected_sheets.items() if selected]

    if not selected_sheet_names:
        add_log("Error: No sheets selected")
        return csv_files

    for sheet_name in selected_sheet_names:
        try:
            # Read the sheet
            df = pd.read_excel(excel_file, sheet_name=sheet_name)

            # Check if dataframe is empty
            if df.empty:
                add_log(f"The sheet '{sheet_name}' is empty.")
                continue

            # Remove unnamed columns for BU POS sheets
            if sheet_name in st.session_state.bupos_sheets or "BU POS" in sheet_name:
                empty_columns = [col for col in df.columns if str(col).startswith('Unnamed')]
                if empty_columns:
                    df = df.drop(columns=empty_columns)
                    add_log(f"Removed {len(empty_columns)} unnamed columns from '{sheet_name}'")

            # Convert to CSV
            csv_buffer = BytesIO()
            df.to_csv(csv_buffer, index=False, quoting=csv.QUOTE_ALL, encoding='utf-8')
            csv_files[sheet_name + ".csv"] = csv_buffer.getvalue()

            add_log(f"Successfully generated CSV for sheet '{sheet_name}'")

        except Exception as e:
            add_log(f"Error processing sheet '{sheet_name}': {str(e)}")

    if csv_files:
        add_log(f"Total: {len(csv_files)} CSV file(s) generated")

    return csv_files

def create_zip(csv_files):
    """Create a zip file containing all CSV files"""
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for filename, content in csv_files.items():
            zip_file.writestr(filename, content)
    return zip_buffer.getvalue()

# Display Anaplan logo
st.markdown("""
    <div style="text-align: left; padding: 0.5rem 0 1rem 0;">
        <svg width="400" height="80" viewBox="0 0 800 150" xmlns="http://www.w3.org/2000/svg">
            <rect width="800" height="150" fill="#1B3D5F"/>
            <text x="50" y="100" font-family="Arial, sans-serif" font-size="80" font-weight="bold" fill="white">/Anaplan</text>
        </svg>
    </div>
    """, unsafe_allow_html=True)

st.title("CSV Generator from Excel")

# File uploader
uploaded_file = st.file_uploader("Select Excel File", type=['xlsx', 'xlsm', 'xls'])

if uploaded_file is not None:
    st.info(f"Selected File: {uploaded_file.name}")

    # Load available sheets
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        st.session_state.available_sheets = excel_file.sheet_names

        # Sheet selection in compact columns
        st.subheader("Select Sheets")

        # Initialize selected_sheets if needed
        for sheet in st.session_state.available_sheets:
            if sheet not in st.session_state.selected_sheets:
                st.session_state.selected_sheets[sheet] = False

        # Display checkboxes in 5 columns for compact layout
        num_sheets = len(st.session_state.available_sheets)
        cols_per_row = 5

        for i in range(0, num_sheets, cols_per_row):
            cols = st.columns(cols_per_row)
            for j, col in enumerate(cols):
                idx = i + j
                if idx < num_sheets:
                    sheet = st.session_state.available_sheets[idx]
                    with col:
                        st.session_state.selected_sheets[sheet] = st.checkbox(
                            sheet,
                            value=st.session_state.selected_sheets[sheet],
                            key=f"check_{sheet}"
                        )

        # Generate button
        if st.button("Generate CSV Files", type="primary"):
            with st.spinner("Processing..."):
                csv_files = generate_csv_files(uploaded_file)

                if csv_files:
                    st.success(f"Successfully generated {len(csv_files)} CSV file(s)")

                    # Download options in compact layout
                    col1, col2 = st.columns([1, 2])

                    with col1:
                        # Download all as ZIP
                        zip_data = create_zip(csv_files)
                        st.download_button(
                            label="Download All as ZIP",
                            data=zip_data,
                            file_name="csv_files.zip",
                            mime="application/zip"
                        )

                    # Individual downloads
                    st.write("**Individual Files:**")

                    # Display download buttons in 2 columns for compactness
                    file_list = list(csv_files.items())
                    for i in range(0, len(file_list), 2):
                        cols = st.columns(2)
                        for j, col in enumerate(cols):
                            idx = i + j
                            if idx < len(file_list):
                                filename, content = file_list[idx]
                                with col:
                                    st.download_button(
                                        label=f"{filename}",
                                        data=content,
                                        file_name=filename,
                                        mime="text/csv",
                                        key=f"download_{filename}"
                                    )
                else:
                    st.error("No CSV files generated. Check the log below.")

                # Log/Status area (compact)
                if st.session_state.log_messages:
                    with st.expander("Processing Log", expanded=False):
                        log_text = "\n".join(st.session_state.log_messages)
                        st.text(log_text)

    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")

else:
    st.info("Please upload an Excel file to begin")
