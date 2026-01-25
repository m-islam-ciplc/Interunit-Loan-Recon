import re
import pandas as pd
from typing import Any, Dict, List, Optional, Set, Tuple

# Import unmatched tracker
try:
    from ..unmatched_tracker import get_unmatched_tracker
except ImportError:
    from unmatched_tracker import get_unmatched_tracker


MONTHS_FULL = "JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER"

# Per latest requirement: simplest rule:
# - Narration must contain SALARY or REMUNERATION
# - Must contain: "MONTH OF <MMMM>-<YY|YYYY>" or "MONTH OF <MMMM> <YY|YYYY>"
MONTH_OF_MONTHYEAR_PATTERN = rf"\bMONTH\s+OF\s+({MONTHS_FULL})\s*[- ]\s*(20\d{{2}}|\d{{2}})\b"

# Additional supported phrasing (requested):
#   "SALARY OF <MMMM>-<YY|YYYY>" / "REMUNERATION OF <MMMM>-<YY|YYYY>"
SALARY_OF_MONTHYEAR_PATTERN = rf"\b(?:SALARY|REMUNERATION)\s+OF\s+({MONTHS_FULL})\s*[- ]\s*(20\d{{2}}|\d{{2}})\b"

# Festival bonus subset (within Salary matching):
# - Must contain FESTIVAL BONUS
# - Must contain Eid type + year, supporting:
#   "EID UL FITR 2025", "EID UL FITR-2025", "EID-UL-FITR-2025", "EID UL-FITR 2025", etc.
FESTIVAL_BONUS_PATTERN = r"\bFESTIVAL\s+BONUS\b"
EID_YEAR_PATTERN = r"\bEID(?:\s*[- ]\s*UL)?\s*[- ]\s*(FITR|AZHA)\s*[- ]\s*(20\d{2})\b"


def _norm_text(s: str) -> str:
    s = (s or "").upper()
    # Keep only letters/numbers/spaces/hyphen so regexes are predictable
    s = re.sub(r"[^A-Z0-9\- ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _extract_month_years(text: str) -> Set[Tuple[str, str]]:
    """Extract all (MONTH_FULL, YYYY) pairs from supported month-year occurrences."""
    t = _norm_text(text)
    results: Set[Tuple[str, str]] = set()
    for pattern in (MONTH_OF_MONTHYEAR_PATTERN, SALARY_OF_MONTHYEAR_PATTERN):
        for m in re.finditer(pattern, t):
            month = m.group(1).upper()
            year_raw = m.group(2)
            year = year_raw if len(year_raw) == 4 else f"20{year_raw}"
            results.add((month, year))
    return results


def _extract_festival_bonus_keys(text: str) -> Set[Tuple[str, str]]:
    """
    Extract all (EID_TYPE, YYYY) keys for festival bonus.
    Requires FESTIVAL BONUS + EID..(FITR|AZHA)..YYYY in the same block text.
    """
    t = _norm_text(text)
    if re.search(FESTIVAL_BONUS_PATTERN, t) is None:
        return set()
    keys: Set[Tuple[str, str]] = set()
    for m in re.finditer(EID_YEAR_PATTERN, t):
        eid_type = m.group(1).upper()
        year = m.group(2)
        keys.add((eid_type, year))
    return keys


def _extract_block_text(block_rows: List[int], df: pd.DataFrame) -> str:
    """Extract all text from column C (Particulars) for all rows in a block."""
    text_parts: List[str] = []
    for row_idx in block_rows:
        # Column index 2 is Particulars
        val = df.iloc[row_idx, 2]
        if pd.notna(val):
            text_parts.append(str(val).strip())
    return " ".join(text_parts)


def _looks_like_salary(text: str) -> bool:
    """
    Minimal gating per requirement:
      - Salary/Remuneration case:
          - Must contain SALARY or REMUNERATION
          - Must contain "MONTH OF <MMMM>-<YY|YYYY>" OR "SALARY/REMUNERATION OF <MMMM>-<YY|YYYY>"
      - Festival bonus case:
          - Must contain FESTIVAL BONUS
          - Must contain Eid type + year (EID(-UL)-FITR/AZHA + YYYY, separators space/hyphen)
    """
    t = _norm_text(text)
    has_sr_keyword = (("SALARY" in t) or ("REMUNERATION" in t))
    has_sr_date = (re.search(MONTH_OF_MONTHYEAR_PATTERN, t) is not None) or (re.search(SALARY_OF_MONTHYEAR_PATTERN, t) is not None)
    salary_like = has_sr_keyword and has_sr_date
    bonus_like = (re.search(FESTIVAL_BONUS_PATTERN, t) is not None) and (re.search(EID_YEAR_PATTERN, t) is not None)
    return salary_like or bonus_like


class SalaryMatchingLogic:
    """
    Salary matching:
    - Uses universal amount/opposite-type filtering (same as other block-based matchers).
    - Both sides must look like salary and contain:
        - 'PAID TO <name>'
        - 'SALARY'
        - 'APRIL-YYYY' or 'APRIL YYYY'
    - Names must match after normalization.
    """

    def __init__(self, block_identifier):
        self.block_identifier = block_identifier

    def find_potential_matches(self, transactions1, transactions2, blocks1, blocks2, file1_path, file2_path, existing_matches=None):
        print("\n=== SALARY MATCHING LOGIC ===")
        print("1. Filtering blocks by matching amounts and opposite transaction types")
        print("2. Checking Salary/Remuneration + Month-Of OR Festival Bonus + Eid + Year")
        print("3. Matching by Month-Year (salary) or Eid+Year (festival bonus) (no name matching)")

        unmatched_tracker = get_unmatched_tracker()
        matches: List[Dict[str, Any]] = []

        matching_pairs = self.block_identifier.filter_blocks_by_matching_amounts(
            blocks1, blocks2, file1_path, file2_path
        )

        for block1_rows, block2_rows, amount, file1_is_lender in matching_pairs:
            header1 = block1_rows[0]
            header2 = block2_rows[0]

            text1 = _extract_block_text(block1_rows, transactions1)
            text2 = _extract_block_text(block2_rows, transactions2)

            # Gate early unless BOTH look salary/remuneration-like (per requirement)
            if not (_looks_like_salary(text1) and _looks_like_salary(text2)):
                continue

            # Build match keys for both salary/remuneration and festival bonus
            salary_my1 = _extract_month_years(text1)
            salary_my2 = _extract_month_years(text2)
            common_salary_my = salary_my1.intersection(salary_my2)

            bonus_keys1 = _extract_festival_bonus_keys(text1)
            bonus_keys2 = _extract_festival_bonus_keys(text2)
            common_bonus = bonus_keys1.intersection(bonus_keys2)

            if not common_salary_my and not common_bonus:
                reason = (
                    "Salary/Bonus mismatch: no common keys "
                    f"(salary_my_file1={sorted(salary_my1)}, salary_my_file2={sorted(salary_my2)}, "
                    f"bonus_file1={sorted(bonus_keys1)}, bonus_file2={sorted(bonus_keys2)})"
                )
                unmatched_tracker.add_unmatched_reason(header1, reason, 1)
                unmatched_tracker.add_unmatched_reason(header2, reason, 2)
                continue

            # Match found
            unmatched_tracker.mark_as_matched(header1, header2)

            amounts1 = self.block_identifier.get_header_row_amounts(header1, file1_path)
            amounts2 = self.block_identifier.get_header_row_amounts(header2, file2_path)

            if common_bonus:
                eid_type, year = sorted(common_bonus)[0]
                month = f"EID_UL_{eid_type}"
                audit_info = f"Festival Bonus Match - {month} {year}"
            else:
                month, year = sorted(common_salary_my)[0]
                audit_info = f"Salary/Remuneration Match - {month} {year}"

            matches.append({
                "match_id": None,  # sequential assigned later
                "Match_Type": "Salary",
                "File1_Index": header1,
                "File2_Index": header2,
                "Salary_Name": "",
                "Salary_Month": month,
                "Salary_Year": year,
                "File1_Debit": amounts1.get("debit", 0),
                "File1_Credit": amounts1.get("credit", 0),
                "File2_Debit": amounts2.get("debit", 0),
                "File2_Credit": amounts2.get("credit", 0),
                "File1_Amount": amounts1.get("amount", 0),
                "File2_Amount": amounts2.get("amount", 0),
                "Amount": amount,
                "Audit_Info": audit_info,
            })

        print(f"Salary Matching complete. Found {len(matches)} matches.")
        return matches

