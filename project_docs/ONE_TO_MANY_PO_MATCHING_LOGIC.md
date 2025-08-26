# One-to-Many PO Matching Logic Documentation

## Overview
The One-to-Many (Aggregated) PO logic is designed for "Bulk Payments." It handles cases where one unit makes a single large payment (Lender) that covers multiple individual purchase orders (Borrower) in the other unit.

## Core Matching Rules

To trigger an automatic One-to-Many match, the following conditions must be met:

### 1. **Multiple PO References**
- The **Lender's narration** must list **multiple Purchase Order numbers**.
- **Example**: `Payment for PO#30298, PO#30305, and PO#30310`.

### 2. **Sum of Borrower Amounts**
- The system finds all individual transactions in the other file that match those specific PO numbers.
- The **total sum** of these individual "Borrower" transactions must **exactly equal** the single "Lender" payment amount.

### 3. **Validation**
- All individual parts must be **Credits** (if the bulk payment is a **Debit**) or vice versa.
- Every PO listed in the bulk payment must be found and accounted for.

## How it works (Simple Steps)
1. **Identify Bulk**: The system finds Lender transactions with many PO numbers.
2. **Collect Parts**: It gathers all Borrower records that match those PO numbers.
3. **Calculate**: It adds up the amounts of all those Borrower records.
4. **Match**: If the `Sum of Parts` == `Bulk Payment`, it links all of them together under a single Match ID.

## Summary for Non-Programmers
This logic is for **bulk payments**. If a single payment (Lender) mentions three PO numbers in its description, the system finds the three separate transactions (Borrower) for those POs. If their total matches the payment amount exactly, they are all matched together.

---
*Internal Technical Note:*
*Class: `AggregatedPOMatchingLogic`*
*Matching Type: One Lender to Many Borrowers.*
