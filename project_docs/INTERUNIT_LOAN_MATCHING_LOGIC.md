# Interunit Loan Matching Logic Documentation

## Overview
The Interunit Loan matching logic connects transactions based on bank account numbers and their corresponding short codes. This logic is specifically designed for cases where one unit pays from a bank account and the other unit references that specific account in their description.

## Core Matching Rules

To trigger an automatic Interunit match, the following conditions must be met:

### 1. **Bank Account Mapping**
- The Lender's **full bank account name** must be saved in the `bank_account_mapping.py` file along with its "alias" or **Short Code**.
- **Example**: `BRAC Bank Ltd.-STD-1540102870121001` is mapped to short codes like `BBL#1001`.

### 2. **Header or "As Per Details" Match**
- The **Lender** must have the full bank account name in its transaction. 
- This can be in the **Header Row** (first row) or, if the header says `(as per details)`, it can be in any of the **Subsequent Rows** of that block.

### 3. **Short Code in Borrower Narration**
- The **Borrower's narration** must contain the **matching short code** of the Lender's bank account.
- **Example**: If the Lender paid from an account mapped to `MTB#4355`, the Borrower's description must include `MTB#4355`.

### 4. **Identical Amounts & Opposite Roles**
- The **Lender's Debit** amount must **exactly equal** the **Borrower's Credit** amount.
- Transaction roles must be opposite (one side Debit, the other side Credit).

## How it works (Simple Steps)
1. **Identify Lender Account**: The system looks at the Lender's block to see which bank account was used.
2. **Lookup Short Code**: It finds the corresponding short code (e.g., `MDB#0313`) from the internal dictionary.
3. **Check Borrower**: It scans the Borrower's narration for that specific short code.
4. **Confirm**: If the code is found and the amounts match, the link is confirmed.

## Summary for Non-Programmers
In order to trigger a match, the **Lender bank account number** and its corresponding alias (Short Code) must be saved in the system. The **Lender** side must show the full account number, and the **Borrower narration** must have the **matching short code** (e.g., `MDB#0313`). Amounts must be identical and roles must be opposite.

---
*Internal Technical Note:*
*Mapping File: `bank_account_mapping.py`*
*Matching Direction: One-way (Borrower narration contains Lender's code).*
*Priority: Runs after PO matching.*
