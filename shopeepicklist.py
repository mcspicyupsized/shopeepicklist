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
        sa = gspread.service_account(filename=credentials_file)
        sh = sa.open(sheet_name)
        wks = sh.worksheet(worksheet_name)
        return wks
    except Exception as e:
        logger.error(f"Failed to connect to Google Sheet: {e}")
        raise


def process_raw_data(data: List[str]) -> List[List[str]]:
    """Process raw data by splitting based on markers and separators."""
    processed_data = []
    marker_regex = r'\[\d+\]'

    for item in data:
        entries = re.split(marker_regex, item)
        entries = [entry.strip() for entry in entries if entry.strip()]
        for entry in entries:
            parts = [part.strip() for part in entry.split(';') if part]
            processed_data.append(parts)

    return processed_data


def extract_skus(sku_data: List[str]) -> List[str]:
    """Extract SKUs from the raw SKU column data."""
    skus_list = []
    for bare in sku_data:
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


def process_individual_sizes(variations: List[str], quantities: List[str]) -> Tuple[
    List[Tuple[str, str]], List[Tuple[str, str]], List[int]]:
    """Process individual size variations."""
    sku_updates = []
    quantity_updates = []
    rows_to_clear = []
    size_regex = r'\(\d+(?:\.\d+)?\" [xX] \d+(?:\.\d+)?\"\) \d+pc'

    for index, (variation, quantity_str) in enumerate(zip(variations, quantities), start=1):
        size_match = re.findall(size_regex, variation)
        if size_match:
            for match in size_match:
                sku_size = SIZE_TO_SKU.get(match)
                if sku_size:
                    SKU, base_quantity_str = sku_size.split(' x ')
                    base_quantity = int(base_quantity_str)
                    quantity = parse_quantity(quantity_str)

                    total_quantity = base_quantity * quantity
                    formatted_quantity = f"Quantity: {total_quantity}"
                    sku_updates.append((f'E{index}', SKU))
                    quantity_updates.append((f'D{index}', formatted_quantity))
                    logger.debug(f"Processing individual size: {match} → SKU: {SKU}, Quantity: {total_quantity}")

    return sku_updates, quantity_updates, rows_to_clear


def process_bundle_sizes(variations: List[str], quantities: List[str]) -> Tuple[List[List[str]], List[int]]:
    """Process bundle size variations."""
    additional_rows = []
    rows_to_clear = []
    size_regex2 = r'\(\d+ SIZES, \d+s\)'

    for index, (variation, quantity_str) in enumerate(zip(variations, quantities), start=1):
        size_matches = re.findall(size_regex2, variation)
        if size_matches:
            for matches in size_matches:
                skus_quantities = SIZES_TO_SKU.get(matches, "").split(',')
                for sku_quantity in skus_quantities:
                    if not sku_quantity:
                        continue

                    SKU, base_quantity_str = sku_quantity.split(' x ')
                    base_quantity = int(base_quantity_str)
                    quantity = parse_quantity(quantity_str)

                    total_quantity = base_quantity * quantity
                    new_row = [variation, '', '', f"Quantity: {total_quantity}", SKU]
                    additional_rows.append(new_row)

                rows_to_clear.append(index)
                logger.debug(f"Processing bundle size: {matches}")

    return additional_rows, rows_to_clear


def process_special_bundles(variations: List[str], quantities: List[str], skus: List[str]) -> Tuple[
    List[List[str]], List[int]]:
    """Process special bundle products."""
    bundle_rows = []
    rows_to_clear = []

    # Process variation-based bundles
    for index, (variation, quantity) in enumerate(zip(variations, quantities), start=1):
        for bundle_name, bundle_key in [
            ("Variation Name:Boot Polishing Pack", "Boot Polishing Pack"),
            ("Variation Name:Pouch and Stick", "Pouch and Stick")
        ]:
            if variation == bundle_name:
                skus_list = BUNDLE_SKUS[bundle_key].split(',')
                for sku in skus_list:
                    sku_id, _ = sku.strip().split(' ')
                    new_row = [variation, '', '', f"{quantity}", sku_id]
                    bundle_rows.append(new_row)
                rows_to_clear.append(index)

    # Process SKU-based bundles
    for index, (sku, quantity) in enumerate(zip(skus, quantities), start=1):
        if sku == "RCK-FULLSET":
            rck_sku_items = BUNDLE_SKUS["RCK-FULLSET"].split(",")
            for rck_sku in rck_sku_items:
                if rck_sku.strip():
                    rck_id = rck_sku.strip().split(" ")[0]
                    new_row = [sku, '', '', f"{quantity}", rck_id]
                    bundle_rows.append(new_row)
            rows_to_clear.append(index)

    return bundle_rows, rows_to_clear


def apply_batch_updates(worksheet, sku_updates, quantity_updates):
    """Apply batch updates to the worksheet."""
    batch_updates = []

    # Add SKU updates to batch
    for cell, sku in sku_updates:
        row_index = int(cell[1:])
        batch_updates.append({
            'range': f'E{row_index}',
            'values': [[sku]]
        })

    # Add quantity updates to batch
    for cell, quantity in quantity_updates:
        row_index = int(cell[1:])
        batch_updates.append({
            'range': f'D{row_index}',
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
        update_range = f'A{next_row}:E{next_row + len(additional_rows) - 1}'
        worksheet.update(update_range, additional_rows)
        logger.info(f"Added {len(additional_rows)} new rows")
    except Exception as e:
        logger.error(f"Error adding new rows: {e}")


def clear_processed_rows(worksheet, rows_to_clear):
    """Clear rows that have been processed and expanded."""
    if not rows_to_clear:
        return

    try:
        for row_index in rows_to_clear:
            worksheet.update(f'A{row_index}:E{row_index}', [['', '', '', '', '']])
        logger.info(f"Cleared {len(rows_to_clear)} processed rows")
    except Exception as e:
        logger.error(f"Error clearing rows: {e}")



def process_picklist(credentials_file = "gspread/picklist.json", sheet_name="Warehouse Test", worksheet_name="Imported Data2"):
    """Main function to process the picklist."""
    try:
        # Connect to Google Sheet
        worksheet = connect_to_sheet(credentials_file, sheet_name, worksheet_name)
        logger.info(f"Connected to sheet: {sheet_name}, worksheet: {worksheet_name}")

        # Fetch all required data at once
        data = worksheet.col_values(1)  # Product name
        variation_names = worksheet.col_values(2)  # Variation name
        quantity_column = worksheet.col_values(4)  # Quantity
        sku_column = worksheet.col_values(5)  # SKU

        # Process raw data
        processed_data = process_raw_data(data)
        worksheet.update(processed_data)
        logger.info("Raw data processed and updated")

        # Updated code with validation
        skus_list = extract_skus(sku_column)
        if skus_list:  # Only update if the list isn't empty
            update_range = f'E1:E{len(skus_list)}'
            update_values = [[sku] for sku in skus_list]
            worksheet.update(range_name=update_range, values=update_values)

        # Process individual sizes
        sku_updates, quantity_updates, size_rows_to_clear = process_individual_sizes(
            variation_names, quantity_column
        )

        # Apply batch updates for individual sizes
        apply_batch_updates(worksheet, sku_updates, quantity_updates)

        # Process bundle sizes
        bundle_additional_rows, bundle_rows_to_clear = process_bundle_sizes(
            variation_names, quantity_column
        )

        # Process special bundles
        special_bundle_rows, special_bundle_clear = process_special_bundles(
            variation_names, quantity_column, sku_column
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