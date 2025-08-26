# Transaction Block Identification Logic

## Overview
A transaction block is a group of consecutive rows that represent a single financial transaction in the Excel file. Each block gets a unique Match ID and Audit Info for reconciliation purposes.

## Transaction Block Structure

### 1. **Header Row**
- **Column A**: Contains `Date`
- **Column B**: Contains `Particulars`
- **Column F**: Contains `Vch Type` (Bold text)
- **Column G**: Contains `Vch No.` (Regular text - not bold, not italic)

### 2. **Opening Balance (NOT a transaction block)**
- **Column A**: Contains a date value
- **Column B**: Contains `Dr` or `Cr`
- **Column C**: Contains `Opening Balance` (Bold text)
- **Column F**: Empty/blank
- **Column G**: Empty/blank
- **Column H**: Contains debit amount (Bold number) OR **Column I**: Contains credit amount (Bold number)
- **Missing "Entered By :"** - not a complete transaction

### 3. **Transaction Block Start (First Row)**
Row where **ALL** of these conditions are met:
- **Column A**: Contains a date value
- **Column B**: Contains `Dr` or `Cr`
- **Column C**: Contains ledger text (Bold text)
- **Column F**: Contains `Vch Type` (Bold text) - **NOT blank**
- **Column G**: Contains `Vch No.` (Regular text - not bold, not italic) - **NOT blank**
- **Column H**: Contains debit amount (Bold number) - **NOT blank**, OR
- **Column I**: Contains credit amount (Bold number) - **NOT blank**

### 4. **Block Content (Middle Rows)**
- **Column A**: Empty/blank
- **Column B**: Empty/blank
- **Column C**: Contains either:
  - **Ledger entries**: Bold text in Column C (can be single or multiple rows)
         - **Narration**: Italic text (not bold, but italic) in Column C - located immediately above "Entered By :"
- **Column F**: Empty/blank
- **Column G**: Empty/blank
- **Column H**: Empty/blank
- **Column I**: Empty/blank

### 5. **Block End (Last Row)**
- **Column A**: Empty/blank
- **Column B**: Contains `Entered By :`
- **Column C**: Contains the person's name (Bold + Italic text)
- **Column F**: Empty/blank
- **Column G**: Empty/blank
- **Column H**: Empty/blank
- **Column I**: Empty/blank

## Example Structure

```
Row 1: [01/Jan/2025] | [Cr] | [LEDGER INFO - BOLD] | [VCH TYPE - BOLD] | [VCH NO - REGULAR] | [DEBIT - BOLD] | [CREDIT - BOLD]     ← Block starts
Row 2: [Empty] | [Empty] | [LEDGER INFO 2 - BOLD] | [Empty] | [Empty] | [Empty] | [Empty]     ← More ledger info
Row 3: [Empty] | [Empty] | [LEDGER INFO 3 - BOLD] | [Empty] | [Empty] | [Empty] | [Empty]     ← More ledger info
Row 4: [Empty] | [Empty] | [NARRATION - REGULAR TEXT] | [Empty] | [Empty] | [Empty] | [Empty] ← Narration (second-to-last)
Row 5: [Empty] | [Entered By :] | [PERSON NAME - BOLD + ITALIC] | [Empty] | [Empty] | [Empty] | [Empty]     ← Block ends
```

## Key Rules

1. **Header Row**: Look for `Date`, `Particulars`, `Vch Type`, `Vch No.` in Columns A, B, F, G
2. **Opening Balance**: Date + Dr/Cr + "Opening Balance" + Debit/Credit - NOT a transaction block
3. **Block Start**: Date + Dr/Cr + Ledger + Vch Type (Bold) + Vch No. (Regular) + Debit/Credit (Bold)
4. **Block Content**: Only Column C has values (ledger entries in Bold, narration in Regular text)
5. **Block End**: "Entered By :" + Person name (Bold + Italic)
6. **Block Span**: Multiple rows from start to end

## Match ID and Audit Info Placement

### **Match ID Placement:**
- **ALL rows** of the transaction block get the same Match ID
- **Purpose**: Ensures entire transaction is identified consistently

### **Audit Info Placement:**
- **ONLY the second-to-last row** (narration row) gets Audit Info
- **Purpose**: Keeps audit information in the logical description location

## Technical Implementation Details

### **Row Indexing:**
- **DataFrame rows**: Start at index 0
- **Excel rows**: Start at row 1 (metadata rows 1-8, header row 9, data from row 10)
- **Conversion**: DataFrame row + 10 = Excel row

### **Transaction Block Expansion:**
1. **Look BACKWARDS** from LC match row to find block start (Date + Dr/Cr + Vch Type + Vch No. + Debit/Credit)
2. **Look FORWARDS** from block start to find block end ("Entered By :")
3. **Include ALL rows** between start and end

### **Format Detection:**
- **Bold Text**: `cell.font.bold = True`
- **Italic Text**: `cell.font.italic = True`
- **Regular Text**: `not cell.font.bold and not cell.font.italic`
       - **Ledger Identification**: `desc_cell.font.bold = True`
       - **Narration Identification**: `not desc_cell.font.bold and desc_cell.font.italic`
- **Entered By Name**: `desc_cell.font.bold = True and desc_cell.font.italic = True`

## Purpose
This logic ensures that:
1. Each transaction block gets its own unique Match ID
2. All rows within a block are consistently identified
3. Audit information is placed logically in the narration row
4. Transaction boundaries are accurately identified
5. Reconciliation can be performed with confidence

## Important Notes

- **LC Numbers** are extracted ONLY from narration cells (regular text - not bold, not italic)
- **Amounts** are found in Columns H & I (Debit/Credit)
- **Transaction types** are determined by which column has the amount
- **Matching requires ALL criteria**: Amount + Transaction Type + LC Number
- **Vch Type and Vch No. must be present** in the first row of a transaction block
- **Opening Balance rows are NOT transaction blocks** - they lack Vch Type, Vch No., and "Entered By :"
