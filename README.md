# 🧾 Tax Invoice Generator

A professional desktop application for generating GST-compliant tax invoices in Excel format. Built with Python and Tkinter, this tool simplifies the invoice creation process with an intuitive user interface.

## 📋 Table of Contents

- [Features](#features)
- [Use Cases](#use-cases)
- [Methodology](#methodology)
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

## 💼 Use Cases

### 1. **Small Business Invoice Management**
Small enterprises and startups can use this application to generate professional GST-compliant invoices without expensive accounting software. Perfect for solar panel companies, electronics retailers, and trading businesses.

### 2. **Quick Invoice Generation**
Sales teams can quickly create invoices with a few mouse clicks. The saved party feature eliminates the need to re-enter buyer information for repeat customers, saving time during high-volume invoice periods.

### 3. **GST Compliance**
Businesses operating in India can ensure GST compliance with automatic tax calculations (CGST and SGST). All invoices include proper tax breakdowns and are formatted according to GST invoice requirements.

### 4. **Inventory-Based Invoicing**
Distributors and wholesalers can invoice multiple items in a single invoice with different quantities, rates, and tax percentages. The application handles complex multi-item scenarios with ease.

### 5. **Professional Invoice Presentation**
Generate polished Excel invoices with consistent formatting, proper alignment, and professional styling. Share formatted invoices directly with clients without manual formatting.

### 6. **Discount Management**
Apply percentage-based discounts to individual items, which are automatically deducted from the final total. Useful for seasonal promotions, bulk discounts, or special customer offers.

### 7. **Date Tracking**
Maintain accurate records with separate sale date and delivery date fields. Essential for tracking invoice issuance and delivery timelines.

## 🔧 Methodology

### Application Architecture

The Tax Invoice Generator follows a **Model-View-Controller (MVC)-inspired architecture**:

#### **1. User Interface Layer (View)**
- **Tkinter GUI Framework**: Cross-platform desktop application
- **Tabbed Interface**: Organized into logical sections
- **Form Validation**: Real-time input validation and error handling
- **Visual Feedback**: Success/error messages guide user actions

#### **2. Business Logic Layer (Controller)**
- **Data Processing**: Handles calculations and data transformations
- **Item Management**: Add, remove, and manage invoice items
- **Party Management**: Retrieve and manage buyer information
- **Calculation Engine**: Computes taxes, discounts, and totals

#### **3. Data Layer (Model)**
- **In-Memory Data Storage**: Item list stored during session
- **Party Database**: Saved parties for quick selection
- **Excel Generation**: Direct export to professional Excel format

### Calculation Methodology

The application implements a **precise tax and discount calculation system**:

```
Formula: Total = (Quantity × Rate) + (CGST Amount + SGST Amount) - Discount Amount

Where:
  - Item Subtotal = Quantity × Rate
  - CGST Amount = (Item Subtotal × CGST%) / 100
  - SGST Amount = (Item Subtotal × SGST%) / 100
  - Discount Amount = (Item Subtotal × Discount%) / 100
  - Total = Item Subtotal + CGST Amount + SGST Amount - Discount Amount
```

**Example Calculation:**
```
Input:
  Quantity: 2 units
  Rate: ₹100 per unit
  Discount: 5%
  CGST: 2.5%
  SGST: 3%

Calculation:
  Item Subtotal = 2 × 100 = ₹200
  CGST Amount = (200 × 2.5) / 100 = ₹5
  SGST Amount = (200 × 3) / 100 = ₹6
  Discount Amount = (200 × 5) / 100 = ₹10
  Final Total = 200 + 5 + 6 - 10 = ₹201
```

### Data Flow

```
User Input (GUI)
    ↓
Input Validation & Error Handling
    ↓
Calculation Engine
    ├── Item Subtotal Calculation
    ├── Tax Amount Calculation (CGST/SGST)
    ├── Discount Calculation
    └── Final Total Calculation
    ↓
Treeview Display (Real-time feedback)
    ↓
Excel Generation
    ├── Format Headers
    ├── Populate Data Rows
    ├── Apply Styling
    └── Save File
    ↓
Output (Invoice_[PartyName]_[Date].xlsx)
```

### Key Features of the Methodology

#### **Input Validation**
- Ensures all required fields are populated
- Validates numeric inputs (Quantity, Rate, Percentages)
- Checks for positive values and valid ranges
- Provides clear error messages to guide corrections

#### **Real-Time Calculation**
- Calculations happen immediately upon item addition
- Treeview updates instantly with calculated values
- Users can verify amounts before saving
- No post-processing delays

#### **Excel Export Strategy**
- Organizes data into a 10-column professional table
- Applies consistent formatting and styling
- Includes summary rows (Total GST, Grand Total)
- Optimizes page layout for printing (1-page width fit)
- Adds company branding and signatures

#### **Data Persistence**
- Session-based storage for current invoice items
- Permanent storage of saved parties in application code
- Excel files saved with unique timestamps
- No data loss even if application crashes

### Technical Implementation Details

#### **Column Structure in Excel Invoice**
```
Column A: Serial Number
Column B: HSN/SAC Code (optional)
Column C: Description of Goods
Column D: Quantity
Column E: Rate (per unit)
Column F: Subtotal (Qty × Rate)
Column G: Discount (%)
Column H: CGST (%)
Column I: SGST (%)
Column J: Total (Incl. Tax)
```

#### **GUI Input Fields**
```
Row 1: HSN/SAC Code | Description | Quantity | Rate
Row 2: Discount (%) | CGST (%) | SGST (%)
```

#### **Formatting Standards**
- **Font**: Calibri, 11pt for data, 16pt bold for headers
- **Alignment**: Center-aligned for numeric data, left-aligned for descriptions
- **Borders**: All cells contain borders for clarity
- **Colors**: Gray header background (#D3D3D3) for distinction
- **Page Setup**: Letter size, horizontally centered, fit to 1 page width

### Error Handling & Validation

The application implements robust error handling:

1. **Field-Level Validation**: Checks before item is added
2. **Type Validation**: Ensures numeric fields contain numbers
3. **Range Validation**: Prevents negative quantities or rates
4. **Empty Field Detection**: Alerts user to missing information
5. **File System Validation**: Checks write permissions before saving
6. **User Feedback**: Clear, actionable error messages

### Performance Considerations

- **Lightweight GUI**: Minimal memory footprint
- **Efficient Calculations**: O(n) complexity for n items
- **Direct Excel Writing**: No intermediate file conversions
- **Responsive UI**: No blocking operations during item addition
- **Quick Party Lookup**: O(1) lookup time for saved parties

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
- `tkinter` - For GUI design
- `pillow` - For image insertion
- `pyinstaller` - For converting the project into an Executable file

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
