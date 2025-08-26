import re
import pandas as pd

# Settlement ID pattern: Matches ID: 12345, Employee ID: 12345, Employee-ID: 12345, etc.
# This pattern is robust to hyphens or spaces between 'Employee' and 'ID'
SETTLEMENT_ID_PATTERN = r'(?:Employee\s*[-\s]?\s*)?ID\s*[:\-]?\s*(\d{5})'
FINAL_SETTLEMENT_KEYWORD = r'final\s+settlement'

# Import unmatched tracker
try:
    from ..unmatched_tracker import get_unmatched_tracker
except ImportError:
    from unmatched_tracker import get_unmatched_tracker

class FinalSettlementMatchingLogic:
    """Handles the logic for finding 'Final Settlement' matches based on Employee IDs."""
    
    def __init__(self, block_identifier):
        """
        Initialize with a shared TransactionBlockIdentifier instance.
        
        Args:
            block_identifier: Shared instance of TransactionBlockIdentifier
        """
        self.block_identifier = block_identifier
        
    def extract_ids(self, narration):
        """Extract all 5-digit IDs found in the narration."""
        if not narration or not isinstance(narration, str):
            return set()
        
        # Matches ID: 12345, Employee ID: 12345, etc.
        matches = re.findall(SETTLEMENT_ID_PATTERN, narration, re.IGNORECASE)
        return set(matches)

    def has_settlement_keyword(self, narration):
        """Check if narration contains settlement-related keywords."""
        if not narration or not isinstance(narration, str):
            return False
        return bool(re.search(FINAL_SETTLEMENT_KEYWORD, narration, re.IGNORECASE))
        
    def find_potential_matches(self, transactions1, transactions2, blocks1, blocks2, file1_path, file2_path, existing_matches=None):
        """
        Find matches based on Employee IDs and 'final settlement' keywords.
        
        Args:
            transactions1: DataFrame for file 1
            transactions2: DataFrame for file 2
            blocks1: List of block row indices for file 1
            blocks2: List of block row indices for file 2
            file1_path: Path to file 1
            file2_path: Path to file 2
            existing_matches: List of matches already found by other logics (not used here but kept for compatibility)
            
        Returns:
            List of match dictionaries
        """
        print(f"\n=== FINAL SETTLEMENT MATCHING LOGIC ===")
        print(f"1. Filtering blocks by matching amounts and opposite transaction types")
        print(f"2. Extracting Employee IDs (ID: XXXXX or Employee ID: XXXXX)")
        print(f"3. Matching IDs and checking for 'final settlement' keywords")
        
        unmatched_tracker = get_unmatched_tracker()
        matches = []
        
        # 1. Filter blocks by matching amounts and opposite transaction types
        # This is the UNIVERSAL filtering logic
        matching_pairs = self.block_identifier.filter_blocks_by_matching_amounts(
            blocks1, blocks2, file1_path, file2_path
        )
        
        print(f"Checking {len(matching_pairs)} potential amount-matched pairs for settlement identifiers.")
        
        for block1_rows, block2_rows, amount, file1_is_lender in matching_pairs:
            header1 = block1_rows[0]
            header2 = block2_rows[0]
            
            # Extract full narration text from the entire block (searching all rows)
            text1 = self._extract_block_text(block1_rows, transactions1)
            text2 = self._extract_block_text(block2_rows, transactions2)
            
            # 2. Look for Employee IDs (must be 5 digits)
            id1_matches = re.findall(SETTLEMENT_ID_PATTERN, text1, re.IGNORECASE)
            id2_matches = re.findall(SETTLEMENT_ID_PATTERN, text2, re.IGNORECASE)
            
            # Check if both have at least one ID and if any IDs match
            common_ids = set(id1_matches).intersection(set(id2_matches))
            
            if common_ids:
                # 3. Match found!
                matched_id = list(common_ids)[0]
                
                # Check for 'final settlement' keyword
                has_keyword1 = re.search(FINAL_SETTLEMENT_KEYWORD, text1, re.IGNORECASE)
                has_keyword2 = re.search(FINAL_SETTLEMENT_KEYWORD, text2, re.IGNORECASE)
                
                audit_info = f"Settlement Match (ID: {matched_id})"
                if has_keyword1 or has_keyword2:
                    audit_info += " - 'Final Settlement' keyword found"
                
                print(f"  MATCH FOUND: ID {matched_id}, Amount {amount}")
                
                # Mark as matched in tracker
                unmatched_tracker.mark_as_matched(header1, header2)
                
                # Get amounts for the match record
                amounts1 = self.block_identifier.get_header_row_amounts(header1, file1_path)
                amounts2 = self.block_identifier.get_header_row_amounts(header2, file2_path)
                
                matches.append({
                    'match_id': None,  # Sequential ID assigned later
                    'Match_Type': 'Settlement',
                    'File1_Index': header1,
                    'File2_Index': header2,
                    'Employee_ID': matched_id,
                    'File1_Date': transactions1.iloc[header1].iloc[0],
                    'File1_Description': transactions1.iloc[header1].iloc[2],
                    'File1_Debit': amounts1.get('debit', 0),
                    'File1_Credit': amounts1.get('credit', 0),
                    'File2_Date': transactions2.iloc[header2].iloc[0],
                    'File2_Description': transactions2.iloc[header2].iloc[2],
                    'File2_Debit': amounts2.get('debit', 0),
                    'File2_Credit': amounts2.get('credit', 0),
                    'File1_Amount': amounts1.get('amount', 0),
                    'File2_Amount': amounts2.get('amount', 0),
                    'File1_Type': 'Lender' if amounts1.get('is_lender') else 'Borrower',
                    'File2_Type': 'Lender' if amounts2.get('is_lender') else 'Borrower',
                    'Audit_Info': audit_info
                })
            else:
                # Only log reasons if it looked like it could be a settlement
                if id1_matches or id2_matches:
                    reason = f"Settlement mismatch: File 1 IDs {id1_matches}, File 2 IDs {id2_matches}"
                    unmatched_tracker.add_unmatched_reason(header1, reason, 1)
                    unmatched_tracker.add_unmatched_reason(header2, reason, 2)
                    
        print(f"Final Settlement Matching complete. Found {len(matches)} matches.")
        return matches

    def _extract_block_text(self, block_rows, df):
        """Extract all text from column C (Particulars) for all rows in a block."""
        text_parts = []
        for row_idx in block_rows:
            # Column index 2 is Particulars
            val = df.iloc[row_idx, 2]
            if pd.notna(val):
                text_parts.append(str(val).strip())
        return " ".join(text_parts)
