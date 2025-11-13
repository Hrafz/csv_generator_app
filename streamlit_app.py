import streamlit as st
import pandas as pd
import csv
from io import BytesIO
import zipfile

st.set_page_config(
    page_title="CSV Generator from Excel",
    page_icon="📊",
    layout="wide"
)

# Initialize session state
if 'available_sheets' not in st.session_state:
    st.session_state.available_sheets = []
if 'selected_sheets' not in st.session_state:
    st.session_state.selected_sheets = {}
if 'sheet_filters' not in st.session_state:
    st.session_state.sheet_filters = {}
if 'bupos_sheets' not in st.session_state:
    st.session_state.bupos_sheets = []
if 'log_messages' not in st.session_state:
    st.session_state.log_messages = []

def add_log(message):
    """Add message to log"""
    st.session_state.log_messages.append(message)

def get_column_by_name_or_index(columns, identifier):
    """Get column by name or numeric index"""
    if identifier.strip().isdigit():
        index = int(identifier.strip())
        if 0 <= index < len(columns):
            return columns[index]
    return identifier.strip()

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

            # Apply filter if specified
            filter_identifier = st.session_state.sheet_filters.get(sheet_name, "").strip()
            if filter_identifier:
                filter_column = get_column_by_name_or_index(df.columns, filter_identifier)

                if filter_column in df.columns:
                    # Convert to numeric, handling comma decimal separators
                    df[filter_column] = df[filter_column].apply(
                        lambda x: x.replace(',', '.') if isinstance(x, str) else x
                    )
                    df[filter_column] = pd.to_numeric(df[filter_column], errors='coerce')
                    df = df.dropna(subset=[filter_column])
                    add_log(f"Applied filter on column '{filter_column}' for sheet '{sheet_name}'")
                else:
                    add_log(f"Warning: Filter column '{filter_identifier}' not found in sheet '{sheet_name}'. No filtering applied.")

            # Check if dataframe is empty after filtering
            if df.empty:
                add_log(f"The sheet '{sheet_name}' is empty after applying filters.")
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

# Main UI
st.title("CSV Generator from Excel")

# File uploader
uploaded_file = st.file_uploader("Select Excel File", type=['xlsx', 'xlsm', 'xls'])

if uploaded_file is not None:
    st.info(f"Selected File: {uploaded_file.name}")

    # Load available sheets
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        st.session_state.available_sheets = excel_file.sheet_names

        # Display available sheets
        st.subheader("Available Sheets")
        st.text(", ".join(st.session_state.available_sheets))

        # Sheet selection and filters
        st.subheader("Set Sheet Names and Filters")

        # Create two columns for layout
        col1, col2 = st.columns([3, 2])

        with col1:
            st.write("**Select Sheets:**")

        with col2:
            st.write("**Filter (column name or index):**")

        # Initialize selected_sheets if needed
        for sheet in st.session_state.available_sheets:
            if sheet not in st.session_state.selected_sheets:
                st.session_state.selected_sheets[sheet] = False
            if sheet not in st.session_state.sheet_filters:
                st.session_state.sheet_filters[sheet] = ""

        # Display checkboxes and filter inputs for each sheet
        for sheet in st.session_state.available_sheets:
            col1, col2 = st.columns([3, 2])

            with col1:
                st.session_state.selected_sheets[sheet] = st.checkbox(
                    sheet,
                    value=st.session_state.selected_sheets[sheet],
                    key=f"check_{sheet}"
                )

            with col2:
                if st.session_state.selected_sheets[sheet]:
                    st.session_state.sheet_filters[sheet] = st.text_input(
                        "Filter",
                        value=st.session_state.sheet_filters[sheet],
                        key=f"filter_{sheet}",
                        label_visibility="collapsed"
                    )

        # BU POS sheet names input
        st.subheader("BU POS Sheet Configuration")
        bupos_input = st.text_input(
            "Enter BU POS sheet names (comma-separated):",
            value=", ".join(st.session_state.bupos_sheets) if st.session_state.bupos_sheets else "",
            help="These sheets will have 'Unnamed' columns removed automatically"
        )

        if bupos_input:
            st.session_state.bupos_sheets = [name.strip() for name in bupos_input.split(",") if name.strip()]

        # Generate button
        st.divider()

        if st.button("Generate CSV Files", type="primary"):
            with st.spinner("Processing..."):
                csv_files = generate_csv_files(uploaded_file)

                if csv_files:
                    st.success(f"Successfully generated {len(csv_files)} CSV file(s)")

                    # Download options
                    st.subheader("Download Files")

                    col1, col2 = st.columns([1, 3])

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
                    for filename, content in csv_files.items():
                        st.download_button(
                            label=f"Download {filename}",
                            data=content,
                            file_name=filename,
                            mime="text/csv",
                            key=f"download_{filename}"
                        )
                else:
                    st.error("No CSV files generated. Check the log below.")

        # Log/Status area
        if st.session_state.log_messages:
            st.divider()
            st.subheader("Processing Log")
            log_text = "\n".join(st.session_state.log_messages)
            st.text_area("Status", value=log_text, height=200, disabled=True)

    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")

else:
    st.info("Please upload an Excel file to begin")
