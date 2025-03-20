import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import json
import pickle
from datetime import datetime


class WarehouseOrderProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Warehouse Order Processor")
        self.root.geometry("700x700")

        # Data storage
        self.warehouse_sequence = None
        self.sequence_loaded = False
        self.last_updated = None
        self.processed_data = None
        self.warehouse_file_path = None

        # Load saved sequence if exists
        self.data_file = "warehouse_sequence.pkl"
        self.load_saved_sequence()

        # Create the UI
        self.create_ui()

    def load_saved_sequence(self):
        """Load saved warehouse sequence data if it exists"""
        try:
            if os.path.exists(self.data_file):
                with open(self.data_file, 'rb') as f:
                    data = pickle.load(f)
                    self.warehouse_sequence = data.get('sequence')
                    self.last_updated = data.get('last_updated')
                    self.warehouse_file_path = data.get('file_path')
                    if self.warehouse_sequence:
                        self.sequence_loaded = True
        except Exception as e:
            print(f"Error loading saved sequence: {e}")

    def save_sequence(self, sequence, file_path):
        """Save warehouse sequence data for future use"""
        try:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            data = {
                'sequence': sequence,
                'last_updated': current_time,
                'file_path': file_path
            }
            with open(self.data_file, 'wb') as f:
                pickle.dump(data, f)

            self.warehouse_sequence = sequence
            self.last_updated = current_time
            self.warehouse_file_path = file_path
            self.sequence_loaded = True

            # Update UI
            self.update_sequence_status()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save sequence: {e}")

    def create_ui(self):
        """Create the user interface"""
        # Main frame
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Step 1: Warehouse Sequence
        sequence_frame = tk.LabelFrame(main_frame, text="Step 1: Warehouse Sequence", padx=10, pady=10)
        sequence_frame.pack(fill=tk.X, pady=10)

        # Status display
        self.status_frame = tk.Frame(sequence_frame)
        self.status_frame.pack(fill=tk.X, pady=5)

        self.update_sequence_status()

        # Buttons for sequence management
        btn_frame = tk.Frame(sequence_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        load_btn = tk.Button(btn_frame, text="Load Warehouse Sequence", command=self.load_warehouse_sequence)
        load_btn.pack(side=tk.LEFT, padx=5)

        clear_btn = tk.Button(btn_frame, text="Clear Saved Sequence", command=self.clear_saved_sequence)
        clear_btn.pack(side=tk.RIGHT, padx=5)

        # Step 2: Shopee Orders
        orders_frame = tk.LabelFrame(main_frame, text="Step 2: Upload Shopee Orders", padx=10, pady=10)
        orders_frame.pack(fill=tk.X, pady=10)

        upload_btn = tk.Button(orders_frame, text="Upload Shopee Orders", command=self.load_shopee_orders)
        upload_btn.pack(pady=10)

        # Results section
        results_frame = tk.LabelFrame(main_frame, text="Results", padx=10, pady=10)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        copy_all_btn = tk.Button(results_frame, text="Copy All SKUs (In Sequence)", command=self.copy_all_skus)
        copy_all_btn.pack(pady=10)

        # Warehouse sections
        self.warehouse_frames = {}
        self.warehouse_lists = {}

        warehouses = ['04-2098-5F', '04-2098-4F', '03-2140', '03-2142']
        for warehouse in warehouses:
            frame = tk.LabelFrame(results_frame, text=warehouse)
            frame.pack(fill=tk.BOTH, expand=True, pady=5)

            btn = tk.Button(frame, text=f"Copy {warehouse} SKUs",
                            command=lambda w=warehouse: self.copy_warehouse_skus(w))
            btn.pack(pady=5)

            listbox = tk.Listbox(frame, height=5)
            listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

            self.warehouse_frames[warehouse] = frame
            self.warehouse_lists[warehouse] = listbox

    def update_sequence_status(self):
        """Update the sequence status display"""
        # Clear existing widgets
        for widget in self.status_frame.winfo_children():
            widget.destroy()

        if self.sequence_loaded:
            status_label = tk.Label(self.status_frame, text="✓ Warehouse sequence loaded", fg="green")
            status_label.pack(anchor=tk.W)

            if self.last_updated:
                update_label = tk.Label(self.status_frame,
                                        text=f"Last updated: {self.last_updated}",
                                        fg="gray")
                update_label.pack(anchor=tk.W)

            if self.warehouse_file_path:
                path_label = tk.Label(self.status_frame,
                                      text=f"File: {os.path.basename(self.warehouse_file_path)}",
                                      fg="gray")
                path_label.pack(anchor=tk.W)
        else:
            status_label = tk.Label(self.status_frame, text="No sequence loaded", fg="gray")
            status_label.pack(anchor=tk.W)

    def load_warehouse_sequence(self):
        """Load and process warehouse sequence Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Warehouse Sequence Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if not file_path:
            return

        try:
            # Read Excel file
            workbook = pd.ExcelFile(file_path)

            # Create mapping of SKUs to their order and warehouse
            sku_order = {}

            # Process each sheet in the workbook
            for sheet_name in workbook.sheet_names:
                if '04-2098' in sheet_name:
                    df = pd.read_excel(workbook, sheet_name)
                    for index, row in df.iterrows():
                        if row.iloc[0] and isinstance(row.iloc[0], str):
                            sku = str(row.iloc[0]).split(' x ')[0].strip()
                            # Determine floor based on index
                            warehouse = '04-2098-5F' if index < 50 else '04-2098-4F'
                            sku_order[sku] = {
                                'warehouse': warehouse,
                                'order': index
                            }

                elif '03-2140' in sheet_name or '03-2142' in sheet_name:
                    df = pd.read_excel(workbook, sheet_name)
                    for index, row in df.iterrows():
                        if row.iloc[0] and isinstance(row.iloc[0], str):
                            sku = str(row.iloc[0]).split(' x ')[0].strip()
                            sku_order[sku] = {
                                'warehouse': sheet_name.replace('warehouse ', ''),
                                'order': index
                            }

            # Save the sequence for future use
            self.save_sequence(sku_order, file_path)

            messagebox.showinfo("Success", "Warehouse sequence loaded successfully!")

            # If we have processed data, reprocess with new sequence
            if self.processed_data is not None:
                self.process_orders(self.processed_data, sku_order)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load warehouse sequence: {e}")

    def clear_saved_sequence(self):
        """Clear the saved warehouse sequence"""
        if os.path.exists(self.data_file):
            os.remove(self.data_file)

        self.warehouse_sequence = None
        self.sequence_loaded = False
        self.last_updated = None
        self.warehouse_file_path = None
        self.processed_data = None

        # Clear UI
        self.update_sequence_status()
        for warehouse in self.warehouse_lists:
            self.warehouse_lists[warehouse].delete(0, tk.END)

        messagebox.showinfo("Success", "Saved sequence cleared!")

    def load_shopee_orders(self):
        """Load and process Shopee orders Excel file"""
        if not self.sequence_loaded:
            messagebox.showerror("Error", "Please load warehouse sequence first!")
            return

        file_path = filedialog.askopenfilename(
            title="Select Shopee Orders Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if not file_path:
            return

        try:
            # Read Excel file
            df = pd.read_excel(file_path)

            # Find SKU column (usually column C)
            # This is a simplification - you may need to adapt to your file structure
            self.process_orders(df, self.warehouse_sequence)

            messagebox.showinfo("Success", "Shopee orders processed successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to process Shopee orders: {e}")

    def process_orders(self, df, sequence):
        """Process orders using warehouse sequence"""
        # Store original data
        self.processed_data = df

        # Initialize warehouse groups
        warehouse_groups = {
            '04-2098-5F': [],
            '04-2098-4F': [],
            '03-2140': [],
            '03-2142': []
        }

        # Find the index of the SKU column
        sku_col = None
        for col in df.columns:
            if col == 'SKU' or col == 'C':
                sku_col = col
                break

        if not sku_col:
            # Try to find header row and get data after that
            for i, row in df.iterrows():
                for col in row.index:
                    if row[col] == 'SKU':
                        # Found SKU column, now get data below
                        df = df.iloc[i + 1:].reset_index(drop=True)
                        sku_col = col
                        break
                if sku_col:
                    break

        # If still no SKU column, try column C (index 2)
        if not sku_col and len(df.columns) > 2:
            sku_col = df.columns[2]

        # Process each order
        for i, row in df.iterrows():
            if sku_col and i < len(df):
                sku_value = row[sku_col]
                if isinstance(sku_value, str) and sku_value.strip():
                    base_sku = sku_value.split(' ')[0].strip()
                    if base_sku in sequence:
                        order_info = sequence[base_sku]
                        warehouse_groups[order_info['warehouse']].append({
                            'sku': sku_value,
                            'order': order_info['order']
                        })

        # Sort each warehouse group by the original order
        for warehouse in warehouse_groups:
            warehouse_groups[warehouse].sort(key=lambda x: x['order'])

            # Update listbox
            listbox = self.warehouse_lists[warehouse]
            listbox.delete(0, tk.END)
            for item in warehouse_groups[warehouse]:
                listbox.insert(tk.END, item['sku'])

        self.warehouse_groups = warehouse_groups

    def copy_warehouse_skus(self, warehouse):
        """Copy SKUs for a specific warehouse to clipboard"""
        if not hasattr(self, 'warehouse_groups') or warehouse not in self.warehouse_groups:
            messagebox.showerror("Error", "No data to copy!")
            return

        items = self.warehouse_groups[warehouse]
        sku_list = ' '.join(item['sku'] for item in items)

        self.root.clipboard_clear()
        self.root.clipboard_append(sku_list)

        messagebox.showinfo("Success", f"{warehouse} SKUs copied to clipboard!")

    def copy_all_skus(self):
        """Copy all SKUs in sequence to clipboard"""
        if not hasattr(self, 'warehouse_groups'):
            messagebox.showerror("Error", "No data to copy!")
            return

        all_skus = []
        for warehouse in ['04-2098-5F', '04-2098-4F', '03-2140', '03-2142']:
            all_skus.extend(self.warehouse_groups[warehouse])

        all_skus.sort(key=lambda x: x['order'])
        sku_list = ' '.join(item['sku'] for item in all_skus)

        self.root.clipboard_clear()
        self.root.clipboard_append(sku_list)

        messagebox.showinfo("Success", "All SKUs copied to clipboard!")


if __name__ == "__main__":
    root = tk.Tk()
    app = WarehouseOrderProcessor(root)
    root.mainloop()