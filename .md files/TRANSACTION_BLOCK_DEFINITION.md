# Transaction Block Definition - Universal Reference

## What is a Transaction Block?

A **transaction block** is a group of consecutive rows in an Excel file that together represent **one complete financial transaction**. Think of it like a paragraph - multiple lines that belong together.

**Key Point**: All matching logic modules (current and future) must use this same definition to find transaction blocks.

---

## The Structure (Top to Bottom)

### **Row 1: The Header Row** (The Most Important Row)
This is where the transaction **starts** and where you find the **main amounts**.

**What to look for:**
- **Column A**: Date (e.g., "01/Jan/2025") - **MUST be present**
- **Column B**: "Dr" or "Cr" (indicates transaction direction) - **MUST be present**
- **Column F**: Voucher Type (Bold text) - **MUST be present**
- **Column G**: Voucher Number (Regular text, not bold, not italic) - **MUST be present**
- **Column H**: Debit amount (Bold number) - **OR**
- **Column I**: Credit amount (Bold number) - **MUST have either H or I**
- **Column B**: Must NOT contain "Opening Balance" - **MUST NOT be Opening Balance**

**Note**: Column C (ledger account) is typically present but is NOT checked for block identification. The code only verifies the above columns.

**⚠️ IMPORTANT**: 
- The **amounts in Column H or I** on this header row are the **official transaction amounts**
- Always use amounts from this row, never from other rows in the block
- This row is guaranteed to be the **first row** in any transaction block

### **Rows 2 to N-2: The Middle Rows** (Block Content)
These rows contain additional details about the transaction.

**What to look for:**
- **Column A**: Empty/blank
- **Column B**: Empty/blank
- **Column C**: Contains either:
  - **Ledger accounts** (Bold text) - Additional accounts involved
  - **Narration** (Italic text, not bold) - Description of the transaction
- **Columns F, G, H, I**: Empty/blank

**How to identify:**
- **Ledger account**: Text in Column C that is **Bold** (but not italic)
- **Narration**: Text in Column C that is **Italic** (but not bold)

### **Row N-1: The Narration Row** (Second-to-Last)
This is typically where the main narration/description appears.

**What to look for:**
- **Column C**: Narration text (Italic, not bold)
- This is where you'll find: LC numbers, PO numbers, USD amounts, short codes, etc.

### **Row N: The End Row** (Last Row)
This marks the **end** of the transaction block.

**What to look for:**
- **Column A**: Empty/blank
- **Column B**: Contains exactly **"Entered By :"** (with colon)
- **Column C**: Person's name (Bold + Italic text)
- **Columns F, G, H, I**: Empty/blank

---

## Where to Find What Data

### **1. Transaction Amounts**
**Location**: Header Row (first row of block), Columns H and I
- **Column H**: Debit amount (if this is a debit transaction)
- **Column I**: Credit amount (if this is a credit transaction)
- **Format**: Bold numbers
- **Rule**: Always use amounts from the header row, never from middle rows

**Universal Method**: Use `block_identifier.get_block_header_amounts(block_rows, file_path)`
- Returns: `{'debit': value, 'credit': value, 'row': header_row_index}`
- For single header row: Use `block_identifier.get_header_row_amounts(header_row_idx, file_path)`

### **2. Ledger Accounts**
**Location**: Header Row and Middle Rows, Column C
- **Format**: Bold text (not italic)
- **Header Row**: First/main ledger account
- **Middle Rows**: Additional ledger accounts (if any)

**How to identify**: `cell.font.bold = True` and `cell.font.italic = False`

### **3. Narration/Description**
**Location**: Middle Rows (especially second-to-last row), Column C
- **Format**: Italic text (not bold)
- **Contains**: LC numbers, PO numbers, USD amounts, short codes, transaction descriptions

**How to identify**: `cell.font.bold = False` and `cell.font.italic = True`

### **4. Transaction Type (Lender vs Borrower)**
**Location**: Header Row, Columns H and I
- **Lender**: Column H has amount (debit > 0)
- **Borrower**: Column I has amount (credit > 0)

### **5. Date**
**Location**: Header Row, Column A
- The transaction date

### **6. Voucher Information**
**Location**: Header Row, Columns F and G
- **Column F**: Voucher Type (Bold text) - Required for block identification
- **Column G**: Voucher Number (Regular text) - Required for block identification

---

## Universal Rules for All Matching Modules

### **Rule 1: Block Identification**
- A transaction block **starts** when you find **ALL** of these conditions:
  - Date in Column A (real date, not empty/None)
  - "Dr" or "Cr" in Column B (exactly)
  - Vch Type in Column F (Bold text) - **Required**
  - Vch No. in Column G (Regular text, not bold, not italic) - **Required**
  - Amount in Column H (Bold) **OR** Column I (Bold) - **Required**
  - Column B does NOT contain "Opening Balance" - **Required**
- A transaction block **ends** when you find:
  - "Entered By :" in Column B (exactly, with colon)

### **Rule 2: Amount Extraction**
- **ALWAYS** extract amounts from the **header row** (first row of block)
- **NEVER** extract amounts from middle rows
- Use the universal method: `get_block_header_amounts()` or `get_header_row_amounts()`

### **Rule 3: Data Extraction**
- **Ledger accounts**: Look in Column C, Bold text (not italic)
- **Narration**: Look in Column C, Italic text (not bold)
- **Amounts**: Look in Columns H and I, Header Row only
- **Dates**: Look in Column A, Header Row only

### **Rule 4: Block Boundaries**
- The header row is **always** `block_rows[0]` (first element)
- All rows from header to "Entered By :" are part of the block
- Opening Balance rows are **NOT** transaction blocks (they lack Vch Type, Vch No., and "Entered By :")

---

## What is NOT a Transaction Block

### **Opening Balance Rows**
These look similar but are **NOT** transaction blocks:
- Has Date (Column A) + Dr/Cr (Column B) + "Opening Balance" (Column C, Bold)
- **Missing**: Vch Type (Column F is empty) and Vch No. (Column G is empty)
- **Missing**: "Entered By :" (Column B)

**Rule**: Skip these rows when identifying transaction blocks.

---

## Example Transaction Block

```
Row 14 (Header): [01/Jan/2025] | [Cr] | [Brac Bank PLC-CD-A/C-2028701210002 - BOLD] | [Empty] | [Empty] | [JV - BOLD] | [VCH123 - REGULAR] | [Empty] | [75000000 - BOLD]
Row 15:         [Empty] | [Empty] | [Interunit Funs Transfer as Interunit Loan A/C-Steel Unit,MDB#0331 - ITALIC] | [Empty] | [Empty] | [Empty] | [Empty] | [Empty] | [Empty]
Row 16:         [Empty] | [Entered By :] | [ashiq - BOLD+ITALIC] | [Empty] | [Empty] | [Empty] | [Empty] | [Empty] | [Empty]
```

**Analysis:**
- **Header Row**: Row 14 - Contains date, Dr/Cr, ledger account, Vch Type, Vch No., and amount (Credit: 75,000,000)
- **Middle Row**: Row 15 - Contains narration (Italic) with short code "MDB#0331"
- **End Row**: Row 16 - Contains "Entered By :" and person name

**What to extract:**
- **Amount**: 75,000,000 (from Row 14, Column I)
- **Ledger Account**: "Brac Bank PLC-CD-A/C-2028701210002" (from Row 14, Column C)
- **Narration**: "Interunit Funs Transfer as Interunit Loan A/C-Steel Unit,MDB#0331" (from Row 15, Column C)
- **Short Code**: "MDB#0331" (found in narration)
- **Transaction Type**: Borrower (Credit > 0)

---

## Technical Implementation

### **Row Indexing**
- **DataFrame rows**: Start at index 0
- **Excel rows**: Start at row 1
- **Conversion**: DataFrame row index + 10 = Excel row number
  - Example: DataFrame row 4 = Excel row 14

### **Universal Methods Available**

1. **`identify_transaction_blocks(transactions_df, file_path)`**
   - Returns: List of transaction blocks
   - Each block is a list of row indices
   - `block[0]` is always the header row

2. **`get_block_header_amounts(block_rows, file_path)`**
   - Returns: `{'debit': value, 'credit': value, 'row': header_row_index}`
   - Use when you have a full block (list of row indices)

3. **`get_header_row_amounts(header_row_idx, file_path)`**
   - Returns: `{'debit': value, 'credit': value, 'row': header_row_index}`
   - Use when you only have a single header row index

4. **`find_transaction_block_header(description_row_idx, transactions_df)`**
   - Returns: Header row index for a given description row
   - Use when you have a row index and need to find its block header

---

## Summary for Developers

**When building a new matching logic module:**

1. ✅ Use `TransactionBlockIdentifier` to identify blocks
2. ✅ Always extract amounts from the **header row** using universal methods
3. ✅ Look for ledger accounts in Column C (Bold text) - Note: Column C is NOT checked for block identification
4. ✅ Look for narration in Column C (Italic text)
5. ✅ Remember: Header row is always `block_rows[0]`
6. ✅ Remember: Opening Balance rows are NOT transaction blocks

**Exact Block Start Conditions (from code):**
```python
is_block_start = (
    has_real_date and           # Column A: Date (not empty, not 'None')
    has_dr_cr and              # Column B: Exactly "Dr" or "Cr"
    has_vch_type and           # Column F: Vch Type (Bold)
    has_vch_no and              # Column G: Vch No. (Regular, not bold, not italic)
    (has_debit or has_credit) and  # Column H (Bold) OR Column I (Bold)
    is_not_opening_balance     # Column B: Does NOT contain "Opening Balance"
)
```

**Universal Methods to Use:**
- `block_identifier.identify_transaction_blocks()` - Get all blocks
- `block_identifier.get_block_header_amounts()` - Get amounts from block
- `block_identifier.get_header_row_amounts()` - Get amounts from header row
- `block_identifier.find_transaction_block_header()` - Find header for a row

**Never:**
- ❌ Extract amounts from middle rows
- ❌ Create custom block identification logic
- ❌ Treat Opening Balance as a transaction block
- ❌ Assume amounts are in any row other than the header

---

## Purpose

This universal definition ensures:
1. ✅ All matching modules use the same block identification logic
2. ✅ Amounts are always extracted correctly from the header row
3. ✅ Transaction boundaries are consistently identified
4. ✅ Future matching modules follow the same pattern
5. ✅ Reconciliation is accurate and reliable
