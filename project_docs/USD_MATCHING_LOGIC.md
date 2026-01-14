# USD Matching Logic Documentation

## Overview
The USD matching logic connects transactions based on foreign currency (USD) amounts mentioned in the descriptions. This is used for international payments where the base currency (BDT) amounts are matched as usual, but the USD value provides the unique identifier.

## Core Matching Rules

To trigger an automatic USD match, the following conditions must be met:

### 1. **Matching USD Values**
- Both transaction blocks must contain the **same USD amount** in their descriptions.
- **Example**: If both narrations mention `$1,500.00` or `USD 1500`, they will match.
- The system extracts numbers following `USD`, `$`, or `$` prefixes.

### 2. **Identical BDT Amounts**
- Even though we match by USD, the **actual ledger amounts** (in BDT) must also be identical.
- **Lender's Debit** must equal **Borrower's Credit**.

### 3. **Opposite Roles (Lender vs. Borrower)**
- One unit must be the **Lender** (Debit).
- The other unit must be the **Borrower** (Credit).

## How it works (Simple Steps)
1. **Amount Filter**: The system first finds transactions with matching BDT amounts.
2. **USD Extraction**: it looks for USD currency patterns in the descriptions.
3. **Validation**: If the same USD amount is found on both sides and roles are opposite, it confirms the match.

## Summary for Non-Programmers
In order to trigger a match, **both file narrations** must mention the **same USD amount** (e.g., `USD 500`). The ledger amounts must also be identical and the transaction types must be opposite (one Debit, one Credit).

---
*Internal Technical Note:*
*Regex Pattern: `(?:USD|\$)\s*([\d,]+(?:\.\d+)?)\b`*
*Priority: Runs after Settlement matching.*
