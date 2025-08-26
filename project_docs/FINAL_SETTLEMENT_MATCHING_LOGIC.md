# Final Settlement Matching Logic Documentation

## Overview
The Final Settlement logic automatically connects transactions related to employee settlements. It relies on shared Employee ID numbers found in the descriptions of both files.

## Core Matching Rules

To trigger an automatic Settlement match, **all** of the following conditions must be met:

### 1. **Matching 5-Digit IDs**
- Both transaction blocks must contain the **same exact 5-digit number** preceded by `ID:` or `Employee ID:`.
- **Example**: If File 1 says `ID: 11370` and File 2 says `Employee ID: 11370`, they will match.
- The system is smart enough to handle variations like `Employee-ID`, `ID :`, or `ID-`.

### 2. **Identical Amounts**
- The **Lender's Debit** amount must **exactly equal** the **Borrower's Credit** amount.
- **Example**: If one side shows `99,317.00`, the other must show exactly `99,317.00`.

### 3. **Opposite Roles (Lender vs. Borrower)**
- One file must be the **Lender** (Debit).
- The other file must be the **Borrower** (Credit).

## How it works (Simple Steps)
1. **Filter**: The system finds pairs with matching amounts.
2. **Search**: It scans the full description of both transactions for employee ID patterns.
3. **Verify**: It compares the IDs. If they match, it also checks for the keyword `"final settlement"` to add more context to the report.
4. **Link**: The transactions are linked with a shared Match ID.

## Summary for Non-Programmers
In order to trigger a match, **Lender narration** has to have `ID:` followed by 5 digits (e.g., `ID: 12345`). **Borrower narration** has to have the **same exact ID**. One of the narrations may also have the words "final settlement". Amounts must be identical and types must be opposite.

---
*Internal Technical Note:*
*Regex Pattern: `(?:Employee\s*[-\s]?\s*)?ID\s*[:\-]?\s*(\d{5})`*
*Priority: Runs after Interunit matching.*
