# 🧾 Tax Invoice Generator

A professional desktop application for generating GST-compliant tax invoices in Excel format. Built with Python and Tkinter, this tool simplifies the invoice creation process with an intuitive user interface.

## 📋 Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Application Structure](#application-structure)
- [Customization](#customization)
- [Troubleshooting](#troubleshooting)
- [License](#license)

## ✨ Features

- **User-Friendly GUI**: Clean and intuitive interface built with Tkinter
- **Saved Party Database**: Quick access to frequently used buyer information
- **Calendar Date Picker**: Easy date selection with DD-MM-YYYY format
- **Quick Item Entry**: Press Enter to add items instantly
- **Real-Time Calculations**: Automatic total amount calculation for each item
- **Excel Export**: Professional Excel invoices with proper formatting and styling
- **GST Compliance**: Includes all necessary GST details and formatting
- **Item Management**: Add, view, and remove items from the invoice
- **Input Validation**: Ensures all required fields are filled correctly

## 🔧 Prerequisites

Before running the application, ensure you have:

- **Python 3.7 or higher** installed on your system
- **pip** (Python package installer)

## 📦 Installation

### Step 1: Clone or Download the Repository

```bash
git clone https://github.com/Kshitijh/Invoice.git
cd Invoice
```

Or download the ZIP file and extract it to your desired location.

### Step 2: Create a Virtual Environment (Recommended)

```bash
# On Windows
python -m venv .venv
.venv\Scripts\activate

# On macOS/Linux
python3 -m venv .venv
source .venv/bin/activate
```

### Step 3: Install Required Dependencies

```bash
pip install -r requitements.txt
```

**Required packages:**
- `openpyxl` - For Excel file generation and formatting
- `tkcalendar` - For calendar date picker widget

### Step 4: Verify Installation

```bash
python invoice_app.py
```

The application window should open successfully.

## 🚀 Usage

### Starting the Application

Run the application using:

```bash
python invoice_app.py
```

### Creating an Invoice

#### 1. **Enter Buyer Details**

- **Load Saved Party**: Select from dropdown to auto-fill buyer information
  - Choose "(New Party)" to enter new buyer details
  - Pre-configured parties: Koustubh Enterprise, ALPHA_TRADING
  
- **Manual Entry**:
  - Buyer's Name
  - GST No.
  - Party's Address

#### 2. **Set Invoice Dates**

- **Sale Date**: Click the calendar icon to select the invoice date
  - Defaults to today's date
  - Format: DD-MM-YYYY
  
- **Delivery Date**: Click the calendar icon to select the delivery date
  - Format: DD-MM-YYYY

#### 3. **Add Items/Goods**

Enter the following details for each item:

- **Description**: Item name or description (e.g., "Solar Panel 100W")
- **Quantity**: Number of units (e.g., 5)
- **Rate (Incl. Tax)**: Price per unit including tax (e.g., 1500.00)

**To Add an Item:**
- Fill in all three fields
- Press **Enter** on your keyboard (from any field)
- The item will be added to the table below with calculated total
- Fields will automatically clear for the next item

**To Remove an Item:**
- Click on the item in the table to select it
- Click the **"Remove Selected Item"** button

#### 4. **Generate Invoice**

Once all details are entered:
- Click the **"Generate Tax Invoice (Excel)"** button
- The application will validate all inputs
- A success message will appear with the filename
- The Excel file will be saved in the current directory

### Generated Invoice Format

The Excel invoice includes:

- **Company Header**: Your company name, address, and GSTIN
- **Buyer Details**: Bill To section with buyer information
- **Invoice Dates**: Sale date and delivery date
- **Itemized Table**: 
  - Serial number
  - Description of goods
  - Quantity
  - Rate (including tax)
  - Total amount per item
- **Grand Total**: Sum of all items with tax included
- **Professional Formatting**: Borders, fonts, alignment, and styling

### File Naming Convention

Generated invoices are saved as:
```
Invoice_[BUYER_NAME]_[DATE].xlsx
```

Example: `Invoice_KOUSTUBH_2025-11-26.xlsx`

## 📁 Application Structure

```
Invoice/
├── invoice_app.py          # Main application file
├── requitements.txt        # Python dependencies
├── README.md              # This file
└── Generated Invoices/    # (Created automatically)
    └── Invoice_*.xlsx     # Generated invoice files
```

### Key Components

#### 1. **SAVED_PARTIES Dictionary**
Located at the top of `invoice_app.py`, stores frequently used buyer information:

```python
SAVED_PARTIES = {
    "Koustubh Enterprise": {
        "name": "Koustubh's Solars Pvt. Ltd.",
        "gst": "27AABBCC1234Z5",
        "address": "Ichalkaranji, Maharashtra",
    },
    # Add more parties here
}
```

#### 2. **InvoiceGeneratorApp Class**
Main application class containing:
- `_setup_buyer_details_frame()`: Creates buyer information section
- `_setup_items_frame()`: Creates goods/items section
- `_add_item()`: Handles item addition with validation
- `_remove_item()`: Removes selected items
- `_generate_invoice()`: Validates and triggers Excel generation
- `_generate_invoice_excel()`: Creates the Excel file with formatting

## 🛠️ Customization

### Update Company Information

Edit the Excel generation section in `invoice_app.py`:

```python
# Line ~275-277
ws['A1'] = "Your Company Name"
ws['A2'] = "Your Company Address, City, State, PIN"
ws['A3'] = "GSTIN: Your GST Number"
```

### Add New Saved Parties

Edit the `SAVED_PARTIES` dictionary:

```python
SAVED_PARTIES = {
    "YOUR_KEY": {
        "name": "Company Name",
        "gst": "GST Number",
        "address": "Complete Address",
    },
}
```

### Modify Date Format

To change the date format, edit line ~95:

```python
date_pattern='dd-mm-yyyy'  # Change to desired format
```

### Customize Excel Styling

Modify the styles section in `_generate_invoice_excel()`:

```python
heading_font = Font(name='Calibri', size=16, bold=True)
header_font = Font(name='Calibri', size=11, bold=True)
fill_header = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
```

## 🐛 Troubleshooting

### Application Won't Start

**Error: Module not found**
```bash
# Ensure all dependencies are installed
pip install -r requitements.txt
```

**Error: Python not recognized**
```bash
# Make sure Python is added to PATH
# Or use full path: C:\Python3x\python.exe invoice_app.py
```

### Calendar Not Showing

**Error: tkcalendar not found**
```bash
pip install tkcalendar
```

### Excel File Not Generating

**Check:**
- All buyer details are filled
- Both dates are selected
- At least one item is added
- Write permissions in the current directory
- No file with the same name is open in Excel

### Items Not Adding

**Verify:**
- Description field is not empty
- Quantity and Rate are valid positive numbers
- Press Enter after filling the fields

### Date Format Issues

Ensure dates are in DD-MM-YYYY format. The calendar widget handles this automatically.

## 📊 Example Workflow

1. **Start Application**: Run `python invoice_app.py`
2. **Select Party**: Choose "Koustubh Enterprise" from dropdown
3. **Set Dates**: Use calendar to select invoice and delivery dates
4. **Add Items**:
   - Description: "Solar Panel 100W"
   - Quantity: 10
   - Rate: 1500
   - Press Enter
5. **Add More Items**: Repeat step 4 as needed
6. **Generate**: Click "Generate Tax Invoice (Excel)"
7. **Success**: Invoice saved as `Invoice_KOUSTUBH_2025-11-26.xlsx`

## 🔐 Data Privacy

- All data is stored locally
- No internet connection required
- No data is sent to external servers
- Invoices are saved only on your computer

## 🤝 Contributing

To contribute to this project:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📝 Future Enhancements

Potential features for future versions:
- PDF export option
- Database integration for party management
- Tax calculation breakdown (CGST/SGST/IGST)
- Invoice numbering system
- Email integration
- Multi-currency support
- Payment terms section
- Digital signature support

## 📄 License

This project is open-source and available for personal and commercial use.

**Happy Invoicing! 🎉**
