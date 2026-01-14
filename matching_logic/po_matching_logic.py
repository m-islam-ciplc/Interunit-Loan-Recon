import pandas as pd

# Import unmatched tracker
try:
    from ..unmatched_tracker import get_unmatched_tracker
except ImportError:
    from unmatched_tracker import get_unmatched_tracker

# PO Number extraction pattern - Dynamic approach using /PO/ as anchor
# Finds PO blocks that are continuous text with /PO/ in them
# More flexible boundaries to catch PO numbers at sentence edges
# Examples: CIL/C//PO//11/2024, CCEL/Reno//PO///2024/9/191024, G24/PO/2024/9/29505
PO_PATTERN = r'(?:^|\s)([A-Z0-9/]+/PO/[A-Z0-9/]+)(?:\s|$|[,\.])'

# Configuration


class POMatchingLogic:
    """Handles the logic for finding PO number matches between two files."""
    
    def __init__(self, block_identifier):
        """
        Initialize with a shared TransactionBlockIdentifier instance.
        
        Args:
            block_identifier: Shared instance of TransactionBlockIdentifier for consistent transaction block logic
        """
        self.block_identifier = block_identifier
    
    def find_potential_matches(self, transactions1, transactions2, po_numbers1, po_numbers2, file1_path=None, file2_path=None, existing_matches=None, match_id_manager=None):
        """Find potential PO number matches between the two files."""
        # Filter rows with PO numbers
        po_transactions1 = transactions1[po_numbers1.notna()].copy()
        po_transactions2 = transactions2[po_numbers2.notna()].copy()
        
        print(f"\nFile 1: {len(po_transactions1)} transactions with PO numbers")
        print(f"File 2: {len(po_transactions2)} transactions with PO numbers")
        
        # Find matches - SAME LOGIC AS LC: Amount → Entered By → PO Number
        matches = []
        
        # Use shared state if provided, otherwise create new
        if existing_matches is None:
            existing_matches = {}
        if match_id_manager is None:
            try:
                from ..match_id_manager import get_match_id_manager
            except ImportError:
                from match_id_manager import get_match_id_manager
            match_id_manager = get_match_id_manager()
        
        print(f"\n=== PO MATCHING LOGIC ===")
        print(f"1. Check if lender debit and borrower credit amounts are EXACTLY the same")
        print(f"2. Check if PO numbers match between them")
        print(f"3. Only if both match: Assign same Match ID")
        print(f"4. IMPORTANT: All transactions with same PO+Amount get SAME Match ID")
        
        # Initialize unmatched tracker for audit info
        unmatched_tracker = get_unmatched_tracker()
        
        # Use shared state for tracking which combinations have already been matched
        # Key: (PO_Number, Amount), Value: match_id
        
        # Process each transaction in File 1 to find matches in File 2
        for idx1, po1 in enumerate(po_numbers1):
            if not po1:
                continue
                
            print(f"\n--- Processing File 1 Row {idx1} with PO: {po1} ---")
            
            # Find the transaction block header row for this PO in File 1 (with caching)
            block_header1 = self.block_identifier.find_transaction_block_header(idx1, transactions1, file1_path)
            
            # Extract amounts and lender/borrower status using universal method
            if file1_path:
                amounts1 = self.block_identifier.get_header_row_amounts(block_header1, file1_path)
                file1_debit = amounts1.get('debit', 0) if amounts1.get('debit') else 0.0
                file1_credit = amounts1.get('credit', 0) if amounts1.get('credit') else 0.0
                file1_is_lender = amounts1.get('is_lender', False)
                file1_is_borrower = amounts1.get('is_borrower', False)
                file1_amount = amounts1.get('amount', 0.0)
            else:
                # Fallback to DataFrame if file_path not provided
                header_row1 = transactions1.iloc[block_header1]
                file1_debit_str = str(header_row1.iloc[7]) if pd.notna(header_row1.iloc[7]) else '0'
                file1_credit_str = str(header_row1.iloc[8]) if pd.notna(header_row1.iloc[8]) else '0'
                try:
                    file1_debit = float(file1_debit_str.replace(',', '')) if file1_debit_str.replace('.', '').replace(',', '').isdigit() else 0.0
                    file1_credit = float(file1_credit_str.replace(',', '')) if file1_credit_str.replace('.', '').replace(',', '').isdigit() else 0.0
                except (ValueError, TypeError):
                    file1_debit, file1_credit = 0.0, 0.0
                file1_is_lender = file1_debit > 0
                file1_is_borrower = file1_credit > 0
                file1_amount = file1_debit if file1_is_lender else file1_credit
            
            print(f"  File 1: Amount={file1_amount}, Type={'Lender' if file1_is_lender else 'Borrower'}")
            
            # Now look for matches in File 2
            for idx2, po2 in enumerate(po_numbers2):
                if not po2:
                    continue
                    
                print(f"    Checking File 2 Row {idx2} with PO: {po2}")
                
                # Find the transaction block header row for this PO in File 2 (with caching)
                block_header2 = self.block_identifier.find_transaction_block_header(idx2, transactions2, file2_path)
                
                # Extract amounts and lender/borrower status using universal method
                if file2_path:
                    amounts2 = self.block_identifier.get_header_row_amounts(block_header2, file2_path)
                    file2_debit = amounts2.get('debit', 0) if amounts2.get('debit') else 0.0
                    file2_credit = amounts2.get('credit', 0) if amounts2.get('credit') else 0.0
                    file2_is_lender = amounts2.get('is_lender', False)
                    file2_is_borrower = amounts2.get('is_borrower', False)
                    file2_amount = amounts2.get('amount', 0.0)
                else:
                    # Fallback to DataFrame if file_path not provided
                    header_row2 = transactions2.iloc[block_header2]
                    file2_debit_str = str(header_row2.iloc[7]) if pd.notna(header_row2.iloc[7]) else '0'
                    file2_credit_str = str(header_row2.iloc[8]) if pd.notna(header_row2.iloc[8]) else '0'
                    try:
                        file2_debit = float(file2_debit_str.replace(',', '')) if file2_debit_str.replace('.', '').replace(',', '').isdigit() else 0.0
                        file2_credit = float(file2_credit_str.replace(',', '')) if file2_credit_str.replace('.', '').replace(',', '').isdigit() else 0.0
                    except (ValueError, TypeError):
                        file2_debit, file2_credit = 0.0, 0.0
                    file2_is_lender = file2_debit > 0
                    file2_is_borrower = file2_credit > 0
                    file2_amount = file2_debit if file2_is_lender else file2_credit
                
                print(f"      File 2: Amount={file2_amount}, Type={'Lender' if file2_is_lender else 'Borrower'}")
                
                # STEP 1: Check if amounts are EXACTLY the same
                if file1_amount != file2_amount:
                    reason = f"Amounts don't match: {file1_amount} vs {file2_amount}"
                    print(f"       REJECTED: {reason}")
                    unmatched_tracker.add_unmatched_reason(block_header1, reason, 1)
                    unmatched_tracker.add_unmatched_reason(block_header2, reason, 2)
                    continue
                
                print(f"       STEP 1 PASSED: Amounts match exactly")
                
                # STEP 2: Check if transaction types are opposite (one lender, one borrower)
                if not ((file1_is_lender and file2_is_borrower) or (file1_is_borrower and file2_is_lender)):
                    reason = f"Transaction types don't match (both same type: {'Lender' if file1_is_lender else 'Borrower'})"
                    print(f"       REJECTED: {reason}")
                    unmatched_tracker.add_unmatched_reason(block_header1, reason, 1)
                    unmatched_tracker.add_unmatched_reason(block_header2, reason, 2)
                    continue
                
                print(f"       STEP 2 PASSED: Transaction types are opposite")
                
                # STEP 3: Check if PO numbers match
                if po1 != po2:
                    reason = f"PO numbers don't match: '{po1}' vs '{po2}'"
                    print(f"       REJECTED: {reason}")
                    unmatched_tracker.add_unmatched_reason(block_header1, reason, 1)
                    unmatched_tracker.add_unmatched_reason(block_header2, reason, 2)
                    continue
                
                print(f"       STEP 3 PASSED: PO numbers match")
                
                # STEP 4: Check if we already have a match for this combination
                # Create new sequential Match ID (simplified approach)
                context = f"PO_{po1}_File1_Row_{idx1}_File2_Row_{idx2}"
                # Match ID will be assigned later in post-processing
                match_id = None
                print(f"       CREATING new Match ID: {match_id}")
                
                print(f"       ALL CRITERIA MET - PO MATCH FOUND!")
                
                # Mark as matched in tracker
                unmatched_tracker.mark_as_matched(block_header1, block_header2)
                
                # Get header row data for match dictionary
                header_row1 = transactions1.iloc[block_header1]
                header_row2 = transactions2.iloc[block_header2]
                
                # Create the match
                matches.append({
                    'match_id': match_id,
                    'Match_Type': 'PO',  # Add explicit match type
                    'File1_Index': block_header1,
                    'File2_Index': block_header2,
                    'PO_Number': po1,
                    'File1_Date': header_row1.iloc[0],
                    'File1_Description': header_row1.iloc[2],
                    'File1_Debit': file1_debit,
                    'File1_Credit': file1_credit,
                    'File2_Date': header_row2.iloc[0],
                    'File2_Description': header_row2.iloc[2],
                    'File2_Debit': file2_debit,
                    'File2_Credit': file2_credit,
                    'File1_Amount': file1_amount,
                    'File2_Amount': file2_amount,
                    'File1_Type': 'Lender' if file1_is_lender else 'Borrower',
                    'File2_Type': 'Lender' if file2_is_lender else 'Borrower'
                })
        
        print(f"\n=== PO MATCHING RESULTS ===")
        print(f"Found {len(matches)} valid PO matches across {len(existing_matches)} unique Match ID combinations!")
        
        # Show some examples
        if matches:
            print("\n=== SAMPLE PO MATCHES ===")
            for i, match in enumerate(matches[:5]):
                print(f"\nPO Match {i+1}:")
                print(f"Match ID: {match['match_id']}")
                print(f"PO Number: {match['PO_Number']}")
                print(f"Amount: {match['File1_Amount']}")
                print(f"File 1: {match['File1_Date']} - {str(match['File1_Description'])[:50]}...")
                print(f"  Type: {match['File1_Type']}, Debit: {match['File1_Debit']}, Credit: {match['File1_Credit']}")
                print(f"File 2: {match['File2_Date']} - {str(match['File2_Description'])[:50]}...")
                print(f"  Type: {match['File2_Type']}, Debit: {match['File2_Debit']}, Credit: {match['File2_Credit']}")
        
        return matches
    
    # Transaction block identification methods are now provided by the shared TransactionBlockIdentifier instance
    # This ensures consistent behavior across all matching modules
    

    

