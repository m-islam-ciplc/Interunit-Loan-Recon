# Salary Matching Logic Documentation

## Overview
The **Salary** matching logic is an additional automatic matching module that runs after **Interunit Matching** and before **Final Settlement** and **USD** matching.

It is intentionally designed as a **keyword + period (month/year or Eid/year) matcher** and does **not** require hardcoded employee/unit names.

## Core Matching Rules

For a Salary match to be created, **all** of the following must be true:

### 1) Identical Amounts + Opposite Roles (Universal Rule)
- The system first filters candidate block pairs using the existing universal rule:
  - **Lender Debit** must equal **Borrower Credit** exactly
  - One side must be Debit (Lender) and the other Credit (Borrower)

### 2) Both Sides Must Contain a Recognizable Salary/Bonus Period Key
Salary matching supports **two key types** (either may trigger a match):

#### A) Salary / Remuneration Key
Both narrations (block text) must contain:
- The keyword **`SALARY`** or **`REMUNERATION`**
- And a period in one of these forms:
  - `MONTH OF <FullMonth>-<YY|YYYY>` (hyphen or space)
  - `SALARY OF <FullMonth>-<YY|YYYY>`
  - `REMUNERATION OF <FullMonth>-<YY|YYYY>`

Examples:
- `... Salary for the month of January-2025 ...`
- `... Salary of January-25 ...`
- `... Remuneration of February 2025 ...`

#### B) Festival Bonus (Subset Within Salary)
Both narrations (block text) must contain:
- The keyword **`FESTIVAL BONUS`**
- And an Eid period key:
  - `EID(-UL)-FITR <YYYY>` or `EID(-UL)-FITR-<YYYY>`
  - `EID(-UL)-AZHA <YYYY>` or `EID(-UL)-AZHA-<YYYY>`
  - Supports mixed hyphen/space separators like `EID UL-FITR 2025`, `EID-UL-FITR-2025`, etc.

Examples:
- `Amount paid to Festival bonus ... agst Eid Ul Fitr-2025 ...`
- `Amount paid to ... festival bonus for Eid ul Azha 2025 ...`

## Output Fields
Matches produced by this logic use:
- `Match_Type`: `Salary`
- `Salary_Month`: Month name for salary/remuneration matches, or `EID_UL_FITR` / `EID_UL_AZHA` for festival bonus
- `Salary_Year`: Normalized to `YYYY` (e.g., `25` becomes `2025`)
- `Audit_Info`: Indicates whether the match came from Salary/Remuneration or Festival Bonus.

## Notes / Guardrails
- No fuzzy name matching is used (to avoid false positives and hardcoding).
- This logic relies on **consistent keyword + period formatting** in narrations.

