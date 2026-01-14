# Narration Matching Logic Documentation

## Overview
The Narration matching logic is the **highest priority** automated matching rule. It connects transactions based on one file's Voucher Number appearing in the other file's description. This is considered a "direct reference" match.

## Core Matching Rules

To trigger an automatic Narration match, the following conditions must be met:

### 1. **Voucher Number Reference**
- One file's **Voucher Number** must be written inside the **other file's narration**.
- **Example**: If File 1 has Voucher No. `70625`, the system looks for the text `"70625"` anywhere in the File 2 descriptions.
- This works both ways: File 1 Vch No in File 2 Narration, OR File 2 Vch No in File 1 Narration.

### 2. **Identical Amounts**
- The **Lender's Debit** amount must **exactly equal** the **Borrower's Credit** amount.
- No cent differences allowed.

### 3. **Opposite Roles (Lender vs. Borrower)**
- One file must be the **Lender** (Debit).
- The other file must be the **Borrower** (Credit).

## How it works (Simple Steps)
1. **Identify Vouchers**: The system picks up the Voucher Number from the first row of a transaction.
2. **Cross-Search**: it searches for that number in the descriptions of all transactions in the other file that have the same amount.
3. **Link**: If found, it immediately confirms the match.

## Summary for Non-Programmers
In order to trigger a match, **one file's Voucher Number** (e.g., `70625`) must be mentioned in the **other file's description**. The amounts must be identical and one must be a Debit while the other is a Credit. Because this is a direct reference, it is the first rule the system checks.

---
*Internal Technical Note:*
*Priority: 1 (Highest)*
*Search Scope: Full block text vs. Opposite Voucher Number.*
