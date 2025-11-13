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
if 'preview_sheet' not in st.session_state:
    st.session_state.preview_sheet = None
if 'current_file' not in st.session_state:
    st.session_state.current_file = None

def add_log(message):
    """Add message to log"""
    st.session_state.log_messages.append(message)

@st.cache_data
def load_sheet_preview(file_bytes, sheet_name, nrows=20):
    """Load preview of a sheet with caching"""
    excel_file = BytesIO(file_bytes)
    df = pd.read_excel(excel_file, sheet_name=sheet_name, nrows=nrows)
    return df

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
            df.to_csv(csv_buffer, index=False, quoting=csv.QUOTE_ALL, encoding='utf-8', sep=';')
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

# Header with logo on the right
col1, col2 = st.columns([4, 1], gap="medium")
with col1:
    st.title("CSV Generator from Excel")
with col2:
    try:
        st.image("assets/anaplan_logo.jpg", use_container_width=True)
    except:
        st.write("")  # If logo not found, skip

# File uploader
uploaded_file = st.file_uploader("Select Excel File", type=['xlsx', 'xlsm', 'xls'])

if uploaded_file is not None:
    st.info(f"Selected File: {uploaded_file.name}")

    # Check if file changed and reset checkboxes if so
    if st.session_state.current_file != uploaded_file.name:
        st.session_state.current_file = uploaded_file.name
        st.session_state.selected_sheets = {}

    # Load available sheets
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        st.session_state.available_sheets = excel_file.sheet_names

        # Sheet selection with generate button on the right
        col_left, col_right = st.columns([4, 1], gap="medium")

        with col_left:
            st.subheader("Select Sheets")

        with col_right:
            generate_button = st.button("Generate CSV Files", type="primary", use_container_width=True)

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

        # Process generate button click
        if generate_button:
            with st.spinner("Processing..."):
                csv_files = generate_csv_files(uploaded_file)

                if csv_files:
                    st.success(f"Successfully generated {len(csv_files)} CSV file(s)")

                    # Download options in compact layout
                    col1, col2 = st.columns([1, 2])

                    with col1:
                        # Download all as ZIP with Excel filename
                        zip_data = create_zip(csv_files)
                        excel_name = uploaded_file.name.rsplit('.', 1)[0]  # Remove extension
                        st.download_button(
                            label="Download All as ZIP",
                            data=zip_data,
                            file_name=f"{excel_name}.zip",
                            mime="application/zip"
                        )

                    # Individual downloads
                    st.write("**Individual Files:**")

                    # Display download buttons in 4 columns for more compact layout
                    file_list = list(csv_files.items())
                    for i in range(0, len(file_list), 4):
                        cols = st.columns(4)
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

        st.divider()

        # Preview section
        st.subheader("Preview Sheet")

        # Dropdown to select sheet for preview
        preview_options = ["None"] + st.session_state.available_sheets
        selected_preview = st.selectbox(
            "Select a sheet to preview:",
            options=preview_options,
            index=0
        )

        # Show preview if selected
        if selected_preview != "None":
            try:
                file_bytes = uploaded_file.getvalue()
                df_preview = load_sheet_preview(file_bytes, selected_preview, nrows=20)

                # Set index to start from 2 (row 1 is headers, like Excel)
                df_preview.index = range(2, len(df_preview) + 2)

                st.write(f"**Preview of '{selected_preview}' (first 20 rows):**")
                st.write(f"Total rows: {len(df_preview)}, Total columns: {len(df_preview.columns)}")
                st.dataframe(df_preview, use_container_width=True)
            except Exception as e:
                st.error(f"Error loading preview: {str(e)}")

    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")

else:
    st.info("Please upload an Excel file to begin")
