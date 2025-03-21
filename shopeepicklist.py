import gspread
import re
import logging
from typing import List, Dict, Tuple, Any

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger('shopee_picklist')

# Configuration dictionaries
SIZE_TO_SKU = {
    '(10" x 7") 5pc': "1152 x 1",
    '(10" x 7") 10pc': "1152 x 2",
    '(10" x 7") 15pc': "1152 x 3",
    '(10" x 7") 20pc': "1152 x 4",
    '(10.5" x 11") 5pc': "1153 x 1",
    '(10.5" x 11") 10pc': "1153 x 2",
    '(10.5" x 11") 15pc': "1153 x 3",
    '(10.5" x 11") 20pc': "1153 x 4",
    '(13" x 16") 5pc': "1154 x 1",
    '(13" x 16") 10pc': "1154 x 2",
    '(13" x 16") 15pc': "1154 x 3",
    '(13" x 16") 20pc': "1154 x 4",
    '(15" X 20") 3pc': "1318 x 1",
    '(15" X 20") 6pc': "1318 x 2",
    '(15" X 20") 9pc': "1318 x 3",
    '(15" X 20") 12pc': "1318 x 4",
}

SIZES_TO_SKU = {
    "(3 SIZES, 5s)": "1152 x 1,1153 x 1,1154 x 1",
    "(3 SIZES, 10s)": "1152 x 2,1153 x 2,1154 x 2",
    "(3 SIZES, 15s)": "1152 x 3,1153 x 3,1154 x 3",
    "(3 SIZES, 20s)": "1152 x 4,1153 x 4,1154 x 4"
}

BUNDLE_SKUS = {
    "Boot Polishing Pack": "941 1, 6425 1",
    "Pouch and Stick": "019 1, 169 1",
    "RCK-FULLSET": "74 1, 924 1, 925 1,926 1,927 1, 940 1, 1124 1"
}


def connect_to_sheet(credentials_file: str, sheet_name: str, worksheet_name: str) -> gspread.Worksheet:
    """Connect to Google Sheet and return worksheet."""
    try:
        logger.info(f"Connecting to Google Sheets with credentials: {credentials_file}")
        if isinstance(credentials_file, dict):
            sa = gspread.service_account_from_dict(credentials_file)
        else:
            sa = gspread.service_account(filename=credentials_file)
        sh = sa.open(sheet_name)
        wks = sh.worksheet(worksheet_name)
        return wks
    except Exception as e:
        logger.error(f"Failed to connect to Google Sheet: {e}")
        raise


def extract_skus(sku_data: List[str]) -> List[str]:
    """Extract SKUs from the raw SKU column data."""
    skus_list = []
    for bare in sku_data:
        if not bare:  # Skip empty cells
            continue
        split_bare = bare.split(':')
        if len(split_bare) > 1:
            skus_list.append(split_bare[1].strip())
        else:
            skus_list.append(bare)

    return skus_list


def parse_quantity(quantity_str: str) -> int:
    """Parse quantity string to an integer."""
    try:
        quantity_str = quantity_str.replace('Quantity: ', '').strip()
        return int(quantity_str)
    except ValueError:
        logger.warning(f"Invalid quantity: {quantity_str}")
        return 0


def process_individual_sizes(order_ids: List[str], quantities: List[str], skus: List[str]) -> Tuple[
    List[Tuple[str, str]], List[Tuple[str, str]], List[int]]:
    """Process individual size variations."""
    sku_updates = []
    quantity_updates = []
    rows_to_clear = []
    size_regex = r'\(\d+(?:\.\d+)?\" [xX] \d+(?:\.\d+)?\"\) \d+pc'

    for index, (order_id, quantity_str, sku) in enumerate(zip(order_ids, quantities, skus), start=1):
        # Skip empty rows
        if not order_id or not quantity_str:
            continue

        # Look for size variations in SKU or order ID
        size_match = None
        for field in [order_id, sku]:
            if field and re.search(size_regex, field):
                size_match = re.search(size_regex, field).group(0)
                break

        if size_match:
            sku_size = SIZE_TO_SKU.get(size_match)
            if sku_size:
                SKU, base_quantity_str = sku_size.split(' x ')
                base_quantity = int(base_quantity_str)
                quantity = parse_quantity(quantity_str)

                total_quantity = base_quantity * quantity
                formatted_quantity = f"Quantity: {total_quantity}"
                # Update in the correct columns - now C for SKU, B for quantity
                sku_updates.append((f'C{index}', SKU))
                quantity_updates.append((f'B{index}', formatted_quantity))
                logger.debug(f"Processing individual size: {size_match} → SKU: {SKU}, Quantity: {total_quantity}")

    return sku_updates, quantity_updates, rows_to_clear


def process_bundle_sizes(order_ids: List[str], quantities: List[str], skus: List[str]) -> Tuple[
    List[List[str]], List[int]]:
    """Process bundle size variations."""
    additional_rows = []
    rows_to_clear = []
    size_regex2 = r'\(\d+ SIZES, \d+s\)'

    for index, (order_id, quantity_str, sku) in enumerate(zip(order_ids, quantities, skus), start=1):
        # Skip empty rows
        if not order_id or not quantity_str:
            continue

        # Look for bundle size variations in SKU or order ID
        size_matches = None
        for field in [order_id, sku]:
            if field and re.search(size_regex2, field):
                size_matches = re.search(size_regex2, field).group(0)
                break

        if size_matches:
            skus_quantities = SIZES_TO_SKU.get(size_matches, "").split(',')
            for sku_quantity in skus_quantities:
                if not sku_quantity:
                    continue

                SKU, base_quantity_str = sku_quantity.split(' x ')
                base_quantity = int(base_quantity_str)
                quantity = parse_quantity(quantity_str)

                total_quantity = base_quantity * quantity
                # Column order: [order_id, quantity, sku, parent_sku]
                new_row = [order_id, f"Quantity: {total_quantity}", SKU, sku]
                additional_rows.append(new_row)

            rows_to_clear.append(index)
            logger.debug(f"Processing bundle size: {size_matches}")

    return additional_rows, rows_to_clear


def process_special_bundles(order_ids: List[str], quantities: List[str], skus: List[str], parent_skus: List[str]) -> \
Tuple[List[List[str]], List[int]]:
    """Process special bundle products."""
    bundle_rows = []
    rows_to_clear = []

    # Process SKU-based bundles
    for index, (order_id, quantity_str, sku, parent_sku) in enumerate(zip(order_ids, quantities, skus, parent_skus),
                                                                      start=1):
        # Skip empty cells
        if not sku or not quantity_str:
            continue

        # Check special bundles
        special_bundle = None

        # Check if SKU is a special bundle
        if sku == "RCK-FULLSET":
            special_bundle = "RCK-FULLSET"
        elif "Boot Polishing Pack" in sku or "Boot Polishing Pack" in order_id:
            special_bundle = "Boot Polishing Pack"
        elif "Pouch and Stick" in sku or "Pouch and Stick" in order_id:
            special_bundle = "Pouch and Stick"

        if special_bundle:
            bundle_items = BUNDLE_SKUS.get(special_bundle, "").split(",")
            for bundle_item in bundle_items:
                if not bundle_item.strip():
                    continue

                sku_id, quantity_multiplier = bundle_item.strip().split(' ')
                quantity = parse_quantity(quantity_str)
                total_quantity = int(quantity_multiplier) * quantity

                # Column order: [order_id, quantity, sku, parent_sku]
                new_row = [order_id, f"Quantity: {total_quantity}", sku_id, parent_sku or sku]
                bundle_rows.append(new_row)

            rows_to_clear.append(index)
            logger.debug(f"Processing special bundle: {special_bundle}")

    return bundle_rows, rows_to_clear


def apply_batch_updates(worksheet, sku_updates, quantity_updates):
    """Apply batch updates to the worksheet."""
    batch_updates = []

    # Add SKU updates to batch
    for cell, sku in sku_updates:
        batch_updates.append({
            'range': cell,
            'values': [[sku]]
        })

    # Add quantity updates to batch
    for cell, quantity in quantity_updates:
        batch_updates.append({
            'range': cell,
            'values': [[quantity]]
        })

    if batch_updates:
        try:
            worksheet.batch_update(batch_updates)
            logger.info(f"Applied {len(batch_updates)} batch updates")
        except Exception as e:
            logger.error(f"Error during batch update: {e}")


def append_new_rows(worksheet, additional_rows):
    """Append new rows to the worksheet."""
    if not additional_rows:
        return

    try:
        next_row = len(worksheet.col_values(1)) + 1
        if next_row > 1 and additional_rows:  # Make sure we have data and a valid row
            update_range = f'A{next_row}:D{next_row + len(additional_rows) - 1}'
            worksheet.update(range_name=update_range, values=additional_rows)
            logger.info(f"Added {len(additional_rows)} new rows")
    except Exception as e:
        logger.error(f"Error adding new rows: {e}")


def clear_processed_rows(worksheet, rows_to_clear):
    """Clear rows that have been processed and expanded."""
    if not rows_to_clear:
        return

    try:
        for row_index in rows_to_clear:
            worksheet.update(f'A{row_index}:D{row_index}', [['', '', '', '']])
        logger.info(f"Cleared {len(rows_to_clear)} processed rows")
    except Exception as e:
        logger.error(f"Error clearing rows: {e}")


def process_picklist(credentials_file="picklist.json", sheet_name="Warehouse Test", worksheet_name="Imported Data2"):
    """Main function to process the picklist."""
    try:
        # Connect to Google Sheet
        worksheet = connect_to_sheet(credentials_file, sheet_name, worksheet_name)
        logger.info(f"Connected to sheet: {sheet_name}, worksheet: {worksheet_name}")

        # Fetch all required data at once - adjusted for new column positions
        order_ids = worksheet.col_values(1)  # Column A: order_sn
        quantities = worksheet.col_values(2)  # Column B: Quantity
        skus = worksheet.col_values(3)  # Column C: SKU Reference No.
        parent_skus = worksheet.col_values(4)  # Column D: Parent SKU Reference No.

        logger.info(f"Fetched {len(order_ids)} rows of data")
        logger.debug(f"Sample data - Orders: {order_ids[:3]}, Quantities: {quantities[:3]}, SKUs: {skus[:3]}")

        if not order_ids or len(order_ids) <= 1:  # Check if we have data (excluding header)
            logger.warning("No data found in the worksheet")
            return False

        # Process individual sizes
        sku_updates, quantity_updates, size_rows_to_clear = process_individual_sizes(
            order_ids, quantities, skus
        )

        # Apply batch updates for individual sizes
        apply_batch_updates(worksheet, sku_updates, quantity_updates)

        # Process bundle sizes
        bundle_additional_rows, bundle_rows_to_clear = process_bundle_sizes(
            order_ids, quantities, skus
        )

        # Process special bundles
        special_bundle_rows, special_bundle_clear = process_special_bundles(
            order_ids, quantities, skus, parent_skus
        )

        # Combine all additional rows and rows to clear
        all_additional_rows = bundle_additional_rows + special_bundle_rows
        all_rows_to_clear = list(set(bundle_rows_to_clear + special_bundle_clear))

        # Append all new rows
        append_new_rows(worksheet, all_additional_rows)

        # Clear all processed rows
        clear_processed_rows(worksheet, all_rows_to_clear)

        logger.info("Picklist processing completed successfully")
        return True

    except Exception as e:
        logger.error(f"Error processing picklist: {e}")
        return False


if __name__ == "__main__":
    process_picklist()