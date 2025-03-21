import pandas as pd
import re
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('data_processor')


def process_shopee_export(file_path):
    """
    Process Shopee export file and convert it to the format needed for picklist processing.

    Args:
        file_path: Path to the Excel file exported from Shopee

    Returns:
        DataFrame with processed data ready for picklist processing
    """
    try:
        # Read the Excel file
        logger.info(f"Reading file: {file_path}")
        df = pd.read_excel(file_path)

        # Check if required columns exist
        required_cols = ['order_sn', 'product_info']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            logger.error(f"Missing required columns: {missing_cols}")
            return None

        # Drop unnecessary columns if they exist
        cols_to_drop = ['remark_from_buyer', 'seller_note', 'tracking_number']
        for col in cols_to_drop:
            if col in df.columns:
                df = df.drop(columns=[col])

        # Split product info by order entries using regex
        logger.info("Splitting product information entries")
        df['split_entries'] = df['product_info'].apply(
            lambda x: re.split(r'(?=\[\d+\])(?<!\[1\])', x)
        )

        # Remove any empty strings that result from splitting
        df['split_entries'] = df['split_entries'].apply(
            lambda lst: [item for item in lst if item]
        )

        # Explode the 'split_entries' list into separate rows
        df = df.explode('split_entries')

        # Clean up each row to remove any unwanted whitespace and newline characters
        df['split_entries'] = df['split_entries'].str.strip()
        df['split_entries'] = df['split_entries'].str.replace(
            r'\r\n', ' ', regex=True
        ).str.replace(
            ' +', ' ', regex=True
        ).str.strip()

        # Extract Quantity
        quantity_pattern = r'Quantity: (\d+);'
        df['Quantity'] = df['split_entries'].str.extract(quantity_pattern, expand=False)

        # Extract SKU Reference No. and Parent SKU Reference No.
        sku_ref_pattern = r'SKU Reference No\.: ([\w&.-]*);'
        parent_sku_ref_pattern = r'Parent SKU Reference No\.: ([\w&.-]*);'

        df['SKU Reference No.'] = df['split_entries'].str.extract(sku_ref_pattern, expand=False)
        df['Parent SKU Reference No.'] = df['split_entries'].str.extract(parent_sku_ref_pattern, expand=False)

        # Special handling for zip lock bags
        specific_product = 'Army NS BMT Reservist NS 10 Pack Zip lock / Food Grade Zip Lock Plastic Bag / Resealable Zip Bag / Clear Storage Bag'
        variation_name_pattern = r'Variation Name:([^;]+);'

        # Extract Variation Name temporarily for specific products
        df['Variation Name'] = df['split_entries'].str.extract(variation_name_pattern, expand=False)

        condition = df['split_entries'].str.contains(specific_product, na=False)
        df.loc[condition, 'SKU Reference No.'] = df.loc[condition, 'Variation Name']
        df.loc[condition, 'Parent SKU Reference No.'] = df.loc[condition, 'Variation Name']

        # Remove the Variation Name column as it's no longer needed
        df = df.drop(columns=["Variation Name"])

        # Clean up extracted data
        # Replace NaN values with empty strings
        df['Quantity'] = df['Quantity'].fillna('1')
        df['SKU Reference No.'] = df['SKU Reference No.'].fillna('')
        df['Parent SKU Reference No.'] = df['Parent SKU Reference No.'].fillna('')

        # Final data formatting for picklist processing
        result_df = df[['order_sn', 'Quantity', 'SKU Reference No.', 'Parent SKU Reference No.']]

        # Format Quantity column for picklist processing
        result_df['Quantity'] = "Quantity: " + result_df['Quantity']

        logger.info(f"Processing complete. Generated {len(result_df)} rows of data.")
        return result_df

    except Exception as e:
        logger.error(f"Error processing file: {e}")
        raise