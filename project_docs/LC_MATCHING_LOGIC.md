# LC Matching Logic Documentation

## Overview
The LC (Letter of Credit) matching logic is implemented in the `LCMatchingLogic` class within `lc_matching_logic.py`. This system performs interunit reconciliation by matching transactions between two Excel files based on specific criteria.

## Core Matching Criteria

### 1. **Amount Matching** (Exact Match - No Tolerance)
- **Lender Debit Amount** (File 1) must **exactly equal** **Borrower Credit Amount** (File 2)
- **No tolerance** is applied - amounts must be identical
- **Column Mapping**: 
  - Debit amounts: Column H (DataFrame index 7)
  - Credit amounts: Column I (DataFrame index 8)

### 2. **Transaction Type Opposition**
- **File 1** must be **Lender** (has Debit amount) while **File 2** must be **Borrower** (has Credit amount)
- **OR** **File 1** must be **Borrower** (has Credit amount) while **File 2** must be **Lender** (has Debit amount)
- **Both files cannot be the same type** (both Lender or both Borrower)

### 3. **LC Number Matching** (Exact Match)
- LC numbers must **exactly match** between both files
- **Extraction Source**: Only from **Narration rows** (italic text, not bold, in Column C)
- **Format**: Regular expression pattern matching for LC number formats

## Matching Process Flow

### **Step-by-Step Validation**
```
1. ✅ Amount Check: file1_amount == file2_amount
2. ✅ Type Check: (file1_is_lender AND file2_is_borrower) OR (file1_is_borrower AND file2_is_lender)
3. ✅ LC Number Check: lc1 == lc2
4. ✅ Match ID Assignment: Create new or reuse existing
```

### **Match ID Generation Logic**
- **Match Key**: `(LC_Number, Amount)`
- **Duplicate Prevention**: If the same combination exists, **reuse** the existing Match ID
- **New Match ID**: If new combination, generate sequential ID (`M001`, `M002`, etc.)
- **Consistency**: Identical transactions across multiple occurrences get the **same Match ID**

## Data Extraction Methods

### **Transaction Block Header Identification**
```python
def find_transaction_block_header(self, description_row_idx, transactions_df):
    # Look backwards from LC row to find block start
    # Block start has: Date + Dr/Cr + Amount (Debit OR Credit)
```



### **Amount and Type Determination**
```python
# File 1 Analysis
file1_debit = header_row1.iloc[7]    # Column H
file1_credit = header_row1.iloc[8]   # Column I
file1_is_lender = file1_debit > 0
file1_is_borrower = file1_credit > 0
file1_amount = file1_debit if file1_is_lender else file1_credit
```

## Output Structure

### **Match Object Properties**
```python
{
    'match_id': 'M001',                    # Unique identifier
    'File1_Index': block_header1,          # File 1 transaction block start
    'File2_Index': block_header2,          # File 2 transaction block start
    'LC_Number': 'L/C-123',                # Matched LC number
    'File1_Date': '2024-01-01',            # File 1 transaction date
    'File1_Description': 'Description',     # File 1 transaction description
    'File1_Debit': 1000.00,                # File 1 debit amount
    'File1_Credit': 0.00,                  # File 1 credit amount
    'File2_Date': '2024-01-01',            # File 2 transaction date
    'File2_Description': 'Description',     # File 2 transaction description
    'File2_Debit': 0.00,                   # File 2 debit amount
    'File2_Credit': 1000.00,               # File 2 credit amount
    'File1_Amount': 1000.00,               # File 1 transaction amount
    'File2_Amount': 1000.00,               # File 2 transaction amount
    'File1_Type': 'Lender',                # File 1 transaction type
    'File2_Type': 'Borrower'               # File 2 transaction type
}
```

## Key Features

### **Duplicate Prevention**
- **Same LC + Amount** = **Same Match ID**
- **No skipped Match IDs** in sequence
- **Consistent identification** across multiple file runs

### **Performance Optimization**
- **Early rejection** at each validation step
- **Efficient block header identification** (backward search)
- **Minimal redundant processing**

### **Audit Trail**
- **Comprehensive logging** of each matching step
- **Clear rejection reasons** for failed matches
- **Sample match display** for verification

## Configuration

### **Amount Tolerance**
```python
AMOUNT_TOLERANCE = 0.01  # Currently set to 0.01 but not used
# All matching requires EXACT amount equality
```

### **Match ID Format**
```python
match_id = f"M{match_counter:03d}"  # M001, M002, M003, etc.
```

## Error Handling

### **Missing Data Scenarios**
- **No LC numbers found**: Returns empty matches list
- **No block header found**: Uses description row as fallback

### **Validation Failures**
- **Amount mismatch**: Logs rejection with specific amounts
- **Type mismatch**: Logs rejection with transaction types
- **LC mismatch**: Logs rejection with both LC numbers

## Usage Example

```python
# Initialize matching logic
matcher = LCMatchingLogic()

# Find potential matches
matches = matcher.find_potential_matches(
    transactions1,      # File 1 DataFrame
    transactions2,      # File 2 DataFrame
    lc_numbers1,        # File 1 LC numbers series
    lc_numbers2         # File 2 LC numbers series
)

# Process results
print(f"Found {len(matches)} matches")
for match in matches:
    print(f"Match {match['match_id']}: {match['LC_Number']}")
```

## Summary

The LC matching logic provides a robust, multi-criteria approach to interunit transaction reconciliation:

1. **Exact amount matching** ensures financial accuracy
2. **Transaction type opposition** validates interunit relationships
3. **LC number matching** provides transaction identification
4. **Duplicate prevention** maintains consistency
5. **Comprehensive logging** enables audit and debugging

This system is designed for high-accuracy reconciliation with clear audit trails and consistent results across multiple executions.
