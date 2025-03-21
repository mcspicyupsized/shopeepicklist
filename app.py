import streamlit as st
import pandas as pd
import io
import tempfile
import os
from datetime import datetime
import gspread
from gspread_dataframe import get_as_dataframe, set_as_dataframe
import json

# Import our modules
from shopeepicklist import process_picklist
from data_processor import process_shopee_export

st.set_page_config(
    page_title="Shopee Picklist Processor",
    page_icon="📦",
    layout="wide"
)

# CSS for better styling
st.markdown("""
<style>
    .main {
        padding: 20px;
    }
    .title-container {
        display: flex;
        align-items: center;
        margin-bottom: 20px;
    }
    .logo {
        margin-right: 10px;
    }
    .subtitle {
        color: #888;
        font-style: italic;
    }
    .success {
        color: green;
        font-weight: bold;
    }
    .error {
        color: red;
        font-weight: bold;
    }
    .info-box {
        background-color: #f0f2f6;
        padding: 15px;
        border-radius: 5px;
        margin-bottom: 20px;
    }
    .step-box {
        background-color: #f8f9fa;
        border-left: 4px solid #4CAF50;
        padding: 10px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# App header
st.markdown("""
<div class="title-container">
    <div class="logo">📦</div>
    <div>
        <h1>Shopee Picklist Processor</h1>
        <p class="subtitle">Convert Shopee orders into warehouse pick items</p>
    </div>
</div>
""", unsafe_allow_html=True)


# Function to get credentials file
def get_credentials_file():
    """Get the path to the credentials file, either from Streamlit secrets or local file"""
    if hasattr(st, "secrets") and "gcp_service_account" in st.secrets:
        # Create a temporary file with the credentials
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json') as f:
            json.dump(st.secrets["gcp_service_account"], f)
            return f.name
    else:
        # Use local file
        return "gspread/picklist.json"  # Update this path if needed


# Sidebar for configuration
st.sidebar.header("Settings")

# Google Sheets Credentials
credentials_option = st.sidebar.radio(
    "Google Sheets Credentials",
    ["Upload JSON", "Use Saved Credentials"]
)

credentials_file = get_credentials_file()

if credentials_option == "Upload JSON":
    uploaded_credentials = st.sidebar.file_uploader(
        "Upload Google Sheets credentials (JSON)",
        type="json"
    )
    if uploaded_credentials:
        # Save uploaded credentials to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.json') as tmp:
            tmp.write(uploaded_credentials.getvalue())
            credentials_file = tmp.name

# Google Sheet details
sheet_name = st.sidebar.text_input("Google Sheet Name", "Warehouse Test")
worksheet_name = st.sidebar.text_input("Worksheet Name", "Imported Data2")

# Main app tabs
tabs = st.tabs(["Data Import", "Process Orders", "Data Preview", "Help"])

# Data Import Tab
with tabs[0]:
    st.header("Import Shopee Order Data")

    st.markdown("""
    <div class="info-box">
        <p>Upload your Shopee order export file to prepare it for processing. 
        The app will convert it to the format needed for the picklist processor.</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload Shopee order export (Excel)", type=["xlsx", "xls"])

    if uploaded_file:
        with st.spinner("Processing order data..."):
            try:
                # Save the uploaded file to a temporary location
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                    tmp.write(uploaded_file.getvalue())
                    temp_file_path = tmp.name

                # Process the file
                processed_df = process_shopee_export(temp_file_path)

                if processed_df is not None:
                    st.success(f"✅ Successfully processed {len(processed_df)} order items!")

                    # Display the processed data
                    st.subheader("Processed Order Data")
                    st.dataframe(processed_df, use_container_width=True)

                    # Option to save directly to Google Sheet
                    if st.button("Upload to Google Sheet", type="primary"):
                        with st.spinner("Uploading to Google Sheet..."):
                            try:
                                # Connect to the Google Sheet
                                sa = gspread.service_account(filename=credentials_file)
                                sh = sa.open(sheet_name)
                                wks = sh.worksheet(worksheet_name)

                                # Clear the current data
                                wks.clear()

                                # Upload the processed data
                                set_as_dataframe(wks, processed_df)

                                st.success("✅ Data uploaded to Google Sheet successfully!")
                            except Exception as e:
                                st.error(f"❌ Error uploading to Google Sheet: {str(e)}")

                    # Option to download as CSV
                    csv = processed_df.to_csv(index=False)
                    st.download_button(
                        label="Download Processed Data as CSV",
                        data=csv,
                        file_name="processed_orders.csv",
                        mime="text/csv"
                    )
                else:
                    st.error("❌ Could not process the file. Please check that it has the required columns.")

                # Clean up the temporary file
                os.unlink(temp_file_path)

            except Exception as e:
                st.error(f"❌ Error processing file: {str(e)}")

# Process Orders Tab
with tabs[1]:
    st.header("Process Orders")

    st.markdown("""
    <div class="info-box">
        <h3>Process Picklist</h3>
        <p>Click the button below to process the Shopee orders in your Google Sheet. 
        The app will:</p>
        <ul>
            <li>Connect to Google Sheets</li>
            <li>Process the order data</li>
            <li>Expand bundled orders into individual SKUs</li>
            <li>Update the sheet with processed data</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

    # Process button
    process_col1, process_col2 = st.columns([1, 3])

    with process_col1:
        process_button = st.button("Process Picklist", type="primary", use_container_width=True)

    with process_col2:
        if process_button:
            with st.spinner("Processing picklist..."):
                try:
                    success = process_picklist(
                        credentials_file=credentials_file,
                        sheet_name=sheet_name,
                        worksheet_name=worksheet_name
                    )
                    if success:
                        st.success("✅ Picklist processed successfully!")
                    else:
                        st.error("❌ Error processing picklist. Check logs for details.")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

    # Run history
    st.subheader("Processing History")

    # In a real app, you'd store this in a database or file
    # Here we're using session state as a simple example
    if 'run_history' not in st.session_state:
        st.session_state.run_history = []

    if process_button and 'success' in locals() and success:
        st.session_state.run_history.append({
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'sheet': sheet_name,
            'worksheet': worksheet_name,
            'status': 'Success'
        })

    # Display history
    if st.session_state.run_history:
        history_df = pd.DataFrame(st.session_state.run_history)
        st.dataframe(history_df, use_container_width=True)
    else:
        st.info("No processing history yet")

# Data Preview Tab
with tabs[2]:
    st.header("Preview Google Sheet Data")

    if st.button("Load Preview Data"):
        try:
            with st.spinner("Loading data preview..."):
                # Connect to the Google Sheet
                sa = gspread.service_account(filename=credentials_file)
                sh = sa.open(sheet_name)
                wks = sh.worksheet(worksheet_name)

                # Get data as DataFrame
                df = get_as_dataframe(wks, evaluate_formulas=True, skiprows=0)
                df = df.dropna(how='all')  # Remove empty rows

                # Show data
                st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.error(f"Error loading preview: {str(e)}")

# Help Tab
with tabs[3]:
    st.header("Help & Instructions")

    st.markdown("""
    ### Complete Workflow

    <div class="step-box">
        <strong>Step 1:</strong> Export your orders from Shopee admin panel as an Excel file
    </div>

    <div class="step-box">
        <strong>Step 2:</strong> Upload the Excel file in the "Data Import" tab
    </div>

    <div class="step-box">
        <strong>Step 3:</strong> Review the processed data and click "Upload to Google Sheet"
    </div>

    <div class="step-box">
        <strong>Step 4:</strong> Go to the "Process Orders" tab and click "Process Picklist"
    </div>

    <div class="step-box">
        <strong>Step 5:</strong> Check the "Data Preview" tab to see the final processed data
    </div>

    ### Data Format Requirements

    Your Google Sheet should have the following columns:
    - Column A: Product names
    - Column B: Variation names
    - Column D: Quantities (format: "Quantity: X")
    - Column E: SKUs

    ### Bundled Products

    The system automatically handles these product types:
    - Individual sizes (e.g., "(10" x 7") 5pc")
    - Bundled sizes (e.g., "(3 SIZES, 5s)")
    - Special bundles: "Boot Polishing Pack", "Pouch and Stick", "RCK-FULLSET"

    ### Need Help?

    If you encounter any issues or need assistance, please contact the support team.
    """, unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown("© 2025 Shopee Warehouse Management")


# Clean up temporary files on app shutdown
def cleanup():
    if credentials_option == "Upload JSON" and 'uploaded_credentials' in locals() and uploaded_credentials:
        try:
            os.unlink(credentials_file)
        except:
            pass


import atexit

atexit.register(cleanup)