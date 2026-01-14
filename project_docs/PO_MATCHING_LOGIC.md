# PO Matching Logic Documentation

## Overview
The PO (Purchase Order) matching logic automatically connects transactions between two files based on Purchase Order numbers. It ensures that payments made by one unit correctly match the receipts or provisions in the other unit.

## Core Matching Rules

To trigger an automatic PO match, **all** of the following conditions must be met:

### 1. **Matching PO Numbers**
- Both transaction blocks must contain the **same Purchase Order number**.
- **Example**: If File 1 says `PO#30298` and File 2 says `PO-30298`, they will match.
- The system scans all rows within a transaction block to find these numbers.

### 2. **Identical Amounts**
- The **Lender's Debit** amount must **exactly equal** the **Borrower's Credit** amount.
- **Example**: If one side shows a payment of `50,000.00`, the other side must show exactly `50,000.00`.
- There is no tolerance for differences—the amounts must be identical to the cent.

### 3. **Opposite Roles (Lender vs. Borrower)**
- One file must be the **Lender** (showing a Debit amount).
- The other file must be the **Borrower** (showing a Credit amount).
- If both files show a Debit or both show a Credit for the same PO, they will **not** match automatically.

## How it works (Simple Steps)
1. **Filter**: The system first finds all pairs of transactions with the same amount.
2. **Scan**: It looks through the narration (descriptions) of both transactions for PO numbers.
3. **Verify**: If the PO numbers match and the transaction roles are opposite, it confirms the match.
4. **Assign**: A unique Match ID (like `M005`) is assigned to both transactions in the output files.

## Summary for Non-Programmers
In order to trigger a match, **Lender narration** must have a PO number (e.g., `PO#12345`) and the **Borrower narration** must have that **same exact PO number**. Additionally, the amounts must be identical and the transaction types must be opposite (one Debit, one Credit).

---
*Internal Technical Note:*
*Extraction Source: Narration rows (Column C)*
*Pattern: `\b(?:P/O|PO)[-#]?\s*(\d{5,})\b`*
*Priority: Runs after Narration and LC matching.*
