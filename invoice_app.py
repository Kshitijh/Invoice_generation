import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from datetime import date, datetime
from tkcalendar import DateEntry
import os

# --- Simulated Database for Saved Parties (from previous script) ---
SAVED_PARTIES = {
    "Koustubh Enterprise": {
        "name": "Koustubh's Solars Pvt. Ltd.",
        "gst": "27AABBCC1234Z5",
        "address": "Ichalkaranji, Maharashtra",
    },
    "ALPHA_TRADING": {
        "name": "Alpha Trading Co.",
        "gst": "09XXYZ1234A1Z9",
        "address": "B-5, Industrial Estate, New Delhi, Delhi",
    },
}

class InvoiceGeneratorApp:
    def __init__(self, master):
        self.master = master
        master.title("🧾 Tax Invoice Generator")
        master.geometry("800x700")
        
        # --- Data storage for Item details ---
        self.items_data = []
        
        # --- Setup the main notebook (tabs) ---
        self.notebook = ttk.Notebook(master)
        self.notebook.pack(pady=10, padx=10, expand=True, fill="both")
        
        self.page1 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.page1, text="Buyer Details & Items")
        
        self._setup_buyer_details_frame(self.page1)
        self._setup_items_frame(self.page1)

    ## ----------------- PAGE 1: BUYER DETAILS -----------------

    def _setup_buyer_details_frame(self, parent):
        """Creates the frame for Buyer and General Invoice Details."""
        frame = ttk.LabelFrame(parent, text="Buyer & Invoice Details")
        frame.pack(fill="x", padx=10, pady=5)
        
        # Grid Configuration for the frame
        for i in range(4):
            frame.columnconfigure(i, weight=1)

        # 1. Saved Party Dropdown
        ttk.Label(frame, text="Load Saved Party:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        party_keys = ["(New Party)"] + list(SAVED_PARTIES.keys())
        self.party_var = tk.StringVar(frame)
        self.party_var.set(party_keys[0]) # Default value
        
        self.party_dropdown = ttk.Combobox(frame, textvariable=self.party_var, values=party_keys, state="readonly")
        self.party_dropdown.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.party_dropdown.bind("<<ComboboxSelected>>", self._load_saved_party)

        # 2. Buyer Details Entry Fields
        self.buyer_entries = {}
        fields = [
            ("Buyer's Name", "name", 1), 
            ("GST No.", "gst", 2), 
            ("Party's Address", "address", 3)
        ]
        
        for i, (label_text, key, row_num) in enumerate(fields):
            ttk.Label(frame, text=f"{label_text}:").grid(row=row_num, column=0, padx=5, pady=2, sticky="w")
            entry = ttk.Entry(frame)
            entry.grid(row=row_num, column=1, padx=5, pady=2, sticky="ew")
            self.buyer_entries[key] = entry

        # 3. Date Fields with Calendar
        self.date_entries = {}
        date_fields = [
            ("Sale Date", "sale_date", 1), 
            ("Delivery Date", "delivery_date", 2)
        ]
        
        # Default dates
        default_sale_date = date.today()
        
        for i, (label_text, key, row_num) in enumerate(date_fields):
            ttk.Label(frame, text=f"{label_text}:").grid(row=row_num, column=2, padx=5, pady=2, sticky="w")
            date_entry = DateEntry(frame, width=18, background='darkblue',
                                  foreground='white', borderwidth=2, 
                                  date_pattern='dd-mm-yyyy')
            date_entry.grid(row=row_num, column=3, padx=5, pady=2, sticky="ew")
            self.date_entries[key] = date_entry
            if key == "sale_date":
                date_entry.set_date(default_sale_date)

    def _load_saved_party(self, event):
        """Loads details from SAVED_PARTIES into entry fields."""
        selected_key = self.party_var.get()
        if selected_key in SAVED_PARTIES:
            details = SAVED_PARTIES[selected_key]
            for key, entry in self.buyer_entries.items():
                entry.delete(0, tk.END)
                entry.insert(0, details.get(key, ""))
        elif selected_key == "(New Party)":
             for entry in self.buyer_entries.values():
                entry.delete(0, tk.END)

    ## ----------------- PAGE 1: ITEMS ENTRY -----------------

    def _setup_items_frame(self, parent):
        """Creates the frame for adding items and the Treeview for display."""
        frame = ttk.LabelFrame(parent, text="Goods Details")
        frame.pack(fill="both", padx=10, pady=5, expand=True)
        
        # 1. Item Entry fields
        entry_frame = ttk.Frame(frame)
        entry_frame.pack(fill="x", padx=5, pady=5)
        
        self.item_entries = {}
        fields = [
            ("Description", "description"), 
            ("Quantity", "quantity"), 
            ("Rate (Incl. Tax)", "rate")
        ]
        
        for i, (label_text, key) in enumerate(fields):
            ttk.Label(entry_frame, text=label_text).grid(row=0, column=i * 2, padx=5, sticky="w")
            entry = ttk.Entry(entry_frame, width=20 if key != "description" else 40)
            entry.grid(row=0, column=i * 2 + 1, padx=5, sticky="w")
            entry.bind("<Return>", lambda event: self._add_item())
            self.item_entries[key] = entry

        # 2. Treeview for displaying added items
        self.tree = ttk.Treeview(frame, columns=("Qty", "Rate", "Total"), show="headings", height=10)
        self.tree.heading("#0", text="Description of Goods")
        self.tree.column("#0", width=300, anchor="w")
        self.tree.heading("Qty", text="Quantity")
        self.tree.column("Qty", width=100, anchor="center")
        self.tree.heading("Rate", text="Rate (Incl. Tax)")
        self.tree.column("Rate", width=150, anchor="e")
        self.tree.heading("Total", text="Total Amount")
        self.tree.column("Total", width=150, anchor="e")
        self.tree.pack(fill="both", padx=5, pady=5, expand=True)

        # Buttons frame at the bottom
        buttons_frame = ttk.Frame(frame)
        buttons_frame.pack(pady=5, anchor="e")
        
        ttk.Button(buttons_frame, text="Remove Selected Item", command=self._remove_item).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Generate Tax Invoice (Excel)", command=self._generate_invoice).pack(side="left", padx=5)
        
    def _add_item(self):
        """Validates input and adds an item to the items_data list and Treeview."""
        try:
            description = self.item_entries["description"].get().strip()
            quantity = float(self.item_entries["quantity"].get())
            rate = float(self.item_entries["rate"].get())
            
            if not description or quantity <= 0 or rate <= 0:
                messagebox.showerror("Input Error", "All fields must be filled, and Quantity/Rate must be positive numbers.")
                return

            total_amount = quantity * rate
            
            item = {
                "description": description,
                "quantity": quantity,
                "rate": rate,
                "total": total_amount,
            }
            self.items_data.append(item)
            
            # Insert into Treeview
            self.tree.insert("", "end", text=description, values=(
                f"{quantity:.2f}", 
                f"{rate:.2f}", 
                f"{total_amount:.2f}"
            ))
            
            # Clear fields after adding
            for key in self.item_entries:
                self.item_entries[key].delete(0, tk.END)
            self.item_entries["description"].focus()
            
        except ValueError:
            messagebox.showerror("Input Error", "Quantity and Rate must be valid numbers.")

    def _remove_item(self):
        """Removes the selected item from the Treeview and the items_data list."""
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Selection Error", "Please select an item to remove.")
            return

        # Get the index of the selected item in the Treeview
        item_id = selected_item[0]
        item_index = self.tree.index(item_id)
        
        # Remove from the Python list
        if 0 <= item_index < len(self.items_data):
            self.items_data.pop(item_index)
            
        # Remove from the Treeview
        self.tree.delete(item_id)

    ## ----------------- INVOICE GENERATION -----------------
    def _get_all_input_data(self):
        """Collects all data from GUI inputs for validation and generation."""
        
        # 1. Buyer Data
        party_details = {
            key: entry.get().strip() for key, entry in self.buyer_entries.items()
        }
        
        # 2. General Invoice Data - Format dates as DD-MM-YYYY
        invoice_details = {
            key: entry.get_date().strftime('%d-%m-%Y') for key, entry in self.date_entries.items()
        }
        
        # 3. Items Data
        items = self.items_data
        
        # --- Basic Validation ---
        if not all(party_details.values()):
            return None, "All Buyer details must be filled out."
        if not all(invoice_details.values()):
            return None, "Both Sale Date and Delivery Date must be filled out."
        if not items:
            return None, "At least one item must be added to the invoice."

        return {
            "party": party_details,
            "invoice": invoice_details,
            "items": items
        }, None

    def _generate_invoice(self):
        """Gathers data, validates, and calls the Excel generation function."""
        data, error = self._get_all_input_data()
        
        if error:
            messagebox.showerror("Generation Error", error)
            return
            
        try:
            # Call the Excel generation function
            filename = self._generate_invoice_excel(data["party"], data["invoice"], data["items"])
            messagebox.showinfo("Success!", f"Invoice successfully generated as:\n**{filename}**\n\nFile saved in the current directory.")
            
        except Exception as e:
            messagebox.showerror("Critical Error", f"An error occurred during file generation: {e}")

    ## ----------------- EXCEL GENERATION LOGIC -----------------
    
    def _generate_invoice_excel(self, party_details, invoice_details, items):
        """Generates the invoice in an Excel file (Adapted from previous script)."""
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Tax Invoice"

        # --- Styles (Basic) ---
        heading_font = Font(name='Calibri', size=16, bold=True)
        header_font = Font(name='Calibri', size=11, bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                        top=Side(style='thin'), bottom=Side(style='thin'))
        fill_header = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

        # --- Invoice Header (Your Company Details) ---
        ws['A1'] = "YOUR COMPANY NAME"
        ws['A1'].font = heading_font
        ws['A2'] = "Your Company Address, City, State"
        ws['A3'] = "GSTIN: Your GST No. (Edit this in the code)"
        
        # --- Buyer Details ---
        ws['A5'] = "BILL TO:"
        ws['A5'].font = header_font
        ws['A6'] = party_details['name']
        ws['A7'] = f"GST No.: {party_details['gst']}"
        ws['A8'] = f"Address: {party_details['address']}"

        # --- Date Details ---
        ws['F5'] = "INVOICE DATE:"
        ws['F5'].font = header_font
        ws['G5'] = invoice_details['sale_date']
        ws['F6'] = "DELIVERY DATE:"
        ws['F6'].font = header_font
        ws['G6'] = invoice_details['delivery_date']
        
        # --- Table Headers (Row 10) ---
        headers = ["#", "Description of Goods", "Quantity", "Rate (Incl. Tax)", "Total Amount"]
        ws.merge_cells('B10:C10') 
        
        col_start = 1
        for i, header in enumerate(headers):
            col = col_start + i
            # Handle the merged column for Description
            cell = ws.cell(row=10, column=col)
            cell.value = header
            cell.font = header_font
            cell.fill = fill_header
            cell.border = border
            if header == "Description of Goods":
                 col_start += 1 

        # --- Table Data ---
        row = 11
        total_invoice_amount = 0
        
        for i, item in enumerate(items):
            
            ws[f'A{row}'] = i + 1
            ws[f'B{row}'] = item['description']
            ws.merge_cells(f'B{row}:C{row}') 
            ws[f'D{row}'] = item['quantity']
            ws[f'E{row}'] = item['rate']
            ws[f'F{row}'] = item['total'] 
            
            # Formatting and border
            for col_idx in range(1, 7):
                 cell = ws.cell(row=row, column=col_idx)
                 cell.border = border
                 if col_idx >= 4: 
                     cell.alignment = Alignment(horizontal='right')
            
            total_invoice_amount += item['total']
            row += 1

        # --- Summary & Totals ---
        ws.merge_cells(start_row=row + 1, start_column=4, end_row=row + 1, end_column=5)
        ws[f'D{row + 1}'] = "TOTAL (Incl. Tax):"
        ws[f'D{row + 1}'].font = header_font
        ws[f'F{row + 1}'] = total_invoice_amount
        ws[f'F{row + 1}'].font = header_font
        ws[f'F{row + 1}'].alignment = Alignment(horizontal='right')
        
        # --- Column Widths for readability ---
        ws.column_dimensions['A'].width = 5 
        ws.column_dimensions['B'].width = 25 
        ws.column_dimensions['C'].width = 1 
        ws.column_dimensions['D'].width = 12 
        ws.column_dimensions['E'].width = 15 
        ws.column_dimensions['F'].width = 15 

        # --- Save the file ---
        # Ensure the filename is safe and unique
        party_name_safe = party_details['name'].split()[0].replace('.', '').upper()
        filename = os.path.join(os.getcwd(), f"Invoice_{party_name_safe}_{date.today()}.xlsx")
        wb.save(filename)
        return filename

# --- Main Tkinter Loop ---
if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceGeneratorApp(root)
    root.mainloop()