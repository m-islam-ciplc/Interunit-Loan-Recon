"""
Manual Matching Logic Module

This module implements the logic for finding potential manual matches:
- Finds unmatched transaction blocks with matching amounts (lender debit = borrower credit)
- Only amount matching, no other criteria
- Returns potential pairs for manual review in GUI
"""

import pandas as pd
from typing import Dict, List, Tuple
from transaction_block_identifier import TransactionBlockIdentifier


class ManualMatchingLogic:
    """Handles the logic for finding potential manual matches based on amount matching only."""
    
    def __init__(self, block_identifier):
        """
        Initialize with a shared TransactionBlockIdentifier instance.
        
        Args:
            block_identifier: Shared instance of TransactionBlockIdentifier for consistent transaction block logic
        """
        self.block_identifier = block_identifier
    
    def find_potential_manual_matches(
        self,
        transactions1: pd.DataFrame,
        transactions2: pd.DataFrame,
        blocks1: List[List[int]],
        blocks2: List[List[int]],
        file1_path: str,
        file2_path: str,
        existing_matches: List[Dict] = None
    ) -> List[Tuple]:
        """
        Find potential manual matches (unmatched blocks with matching amounts).
        
        Args:
            transactions1: DataFrame of transactions from first file
            transactions2: DataFrame of transactions from second file
            blocks1: List of transaction blocks from File 1
            blocks2: List of transaction blocks from File 2
            file1_path: Path to first file
            file2_path: Path to second file
            existing_matches: List of already matched records (to filter out)
            
        Returns:
            List of tuples: [(block1_rows, block2_rows, amount, file1_is_lender), ...]
            Each tuple contains:
            - block1_rows: List of row indices for File 1 block
            - block2_rows: List of row indices for File 2 block
            - amount: Matching amount (lender debit = borrower credit)
            - file1_is_lender: Boolean indicating if File 1 is the lender
        """
        if existing_matches is None:
            existing_matches = []
        
        # Get all matched block header row indices to filter them out
        matched_file1_indices = set()
        matched_file2_indices = set()
        
        for match in existing_matches:
            if 'File1_Index' in match:
                matched_file1_indices.add(match['File1_Index'])
            if 'File2_Index' in match:
                matched_file2_indices.add(match['File2_Index'])
        
        # Filter blocks to only unmatched ones
        unmatched_blocks1 = []
        unmatched_blocks2 = []
        
        for block in blocks1:
            # Check if any row in this block has been matched
            header_row = block[0] if block else None
            if header_row is not None and header_row not in matched_file1_indices:
                unmatched_blocks1.append(block)
        
        for block in blocks2:
            # Check if any row in this block has been matched
            header_row = block[0] if block else None
            if header_row is not None and header_row not in matched_file2_indices:
                unmatched_blocks2.append(block)
        
        print(f"\n=== MANUAL MATCHING LOGIC ===")
        print(f"Finding potential manual matches (amount matching only)...")
        print(f"Unmatched blocks: File 1: {len(unmatched_blocks1)}, File 2: {len(unmatched_blocks2)}")
        
        # Use universal filter method to find pairs with matching amounts and opposite types
        matching_pairs = self.block_identifier.filter_blocks_by_matching_amounts(
            unmatched_blocks1, unmatched_blocks2, file1_path, file2_path
        )
        
        print(f"Found {len(matching_pairs)} potential manual match pairs")
        
        return matching_pairs
