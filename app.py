import streamlit as st
import pandas as pd
import io
import tempfile
import os
from datetime import datetime
from shopeepicklist import process_picklist  # Make sure this matches your actual filename

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
</style>
""", unsafe_allow_html=True)

# App header
st.markdown("""
<div class="title-container">
    <div class="logo">📦</div>
    <div>
        <h1>Shopee Picklist Processor</h1>
        <p class="subtitle">Convert bulk orders into warehouse pick items</p>
    </div>
</div>
""", unsafe_allow_html=True)

# Sidebar for configuration
st.sidebar.header("Settings")

# Google Sheets Credentials
credentials_option = st.sidebar.radio(
    "Google Sheets Credentials",
    ["Upload JSON", "Use Saved Credentials"]
)

credentials_file = "picklist.json"  # Default saved credentials path

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

# Main app functionality
tab1, tab2, tab3 = st.tabs(["Process Orders", "Data Preview", "Help"])

with tab1:
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

with tab2:
    st.subheader("Preview Google Sheet Data")

    if st.button("Load Preview Data"):
        try:
            with st.spinner("Loading data preview..."):
                # Import here to avoid requiring gspread when not used
                import gspread
                from gspread_dataframe import get_as_dataframe

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

with tab3:
    st.subheader("Help & Instructions")

    st.markdown("""
    ### How to Use This App

    1. **Set Up Google Sheets Credentials**
       - Upload your service account JSON file or use saved credentials
       - Make sure your service account has access to the Google Sheet

    2. **Configure Sheet Details**
       - Enter the exact name of your Google Sheet
       - Enter the worksheet name where your data is located

    3. **Process Your Picklist**
       - Click the "Process Picklist" button on the Process Orders tab
       - Wait for the processing to complete

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
    """)

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
#