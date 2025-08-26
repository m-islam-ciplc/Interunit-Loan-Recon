import pandas as pd
import re
from .po_matching_logic import PO_PATTERN

class AggregatedPOMatchingLogic:
    """Handles the logic for finding aggregated PO matches between two files."""
    
    def __init__(self, block_identifier):
        """
        Initialize with a shared TransactionBlockIdentifier instance.
        
        Args:
            block_identifier: Shared instance of TransactionBlockIdentifier for consistent transaction block logic
        """
        self.block_identifier = block_identifier
    
    def find_potential_matches(self, transactions1, transactions2, po_numbers1, po_numbers2, file1_path=None, file2_path=None, existing_matches=None, match_id_manager=None):
        """Find potential aggregated PO matches between the two files."""
        
        print(f"\nFile 1: {len(transactions1)} transactions")
        print(f"File 2: {len(transactions2)} transactions")
        
        # Find matches - Aggregated PO Logic
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
        
        print(f"\n=== AGGREGATED PO MATCHING LOGIC ===")
        print(f"1. Find lender transactions with MULTIPLE PO numbers in narration")
        print(f"2. Find borrower transactions matching ANY of those PO numbers")
        print(f"3. Validate: Lender_Debit_Amount == Sum(All_Matching_Borrower_Credit_Amounts)")
        print(f"4. Ensure ALL lender POs are present in borrower transactions")
        
        # Process each transaction in File 1 to find multi-PO lender transactions
        processed_narrations = set()  # Track which narrations we've already processed
        
        for idx1 in range(len(transactions1)):
            # Skip if we've already processed this narration
            if idx1 in processed_narrations:
                continue
                
            # Find the transaction block header row for this index
            block_header1 = self.block_identifier.find_transaction_block_header(idx1, transactions1)
            
            # Extract amounts and lender/borrower status using universal method
            if file1_path:
                amounts1 = self.block_identifier.get_header_row_amounts(block_header1, file1_path)
                file1_debit = amounts1.get('debit', 0) if amounts1.get('debit') else 0.0
                file1_credit = amounts1.get('credit', 0) if amounts1.get('credit') else 0.0
                file1_is_lender = amounts1.get('is_lender', False)
                file1_is_borrower = amounts1.get('is_borrower', False)
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
            
            # Only process lender transactions (debit > 0)
            if not file1_is_lender:
                continue
                
            # Get header row for narration extraction
            header_row1 = transactions1.iloc[block_header1]
            
            # Extract narration from the transaction block
            narration1 = str(header_row1.iloc[2]).strip()
            
            # Skip empty or very short narrations
            if len(narration1) < 10 or narration1.lower() in ['nan', 'none', '']:
                continue
            
            # Extract ALL PO numbers from this narration
            po_matches1 = re.findall(PO_PATTERN, narration1)
            
            # Only process if there are MULTIPLE POs (aggregated scenario)
            if len(po_matches1) < 2:
                continue
                
            print(f"\n--- Processing File 1 Row {block_header1} (Multi-PO Lender) ---")
            print(f"  Narration: {narration1[:80]}...")
            print(f"  Found {len(po_matches1)} POs: {po_matches1}")
            print(f"  Lender Amount: {file1_debit}")
            
            # Mark this narration as processed
            processed_narrations.add(block_header1)
            
            # Now search for matching borrower transactions in File 2
            matching_borrower_transactions = []
            total_borrower_amount = 0
            found_pos = set()
            
            for idx2 in range(len(transactions2)):
                # Find the transaction block header row for this index in File 2
                block_header2 = self.block_identifier.find_transaction_block_header(idx2, transactions2)
                
                # Extract amounts and lender/borrower status using universal method
                if file2_path:
                    amounts2 = self.block_identifier.get_header_row_amounts(block_header2, file2_path)
                    file2_debit = amounts2.get('debit', 0) if amounts2.get('debit') else 0.0
                    file2_credit = amounts2.get('credit', 0) if amounts2.get('credit') else 0.0
                    file2_is_lender = amounts2.get('is_lender', False)
                    file2_is_borrower = amounts2.get('is_borrower', False)
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
                
                # Only process borrower transactions (credit > 0)
                if not file2_is_borrower:
                    continue
                
                # Get header row for narration extraction
                header_row2 = transactions2.iloc[block_header2]
                
                # Extract narration from the transaction block in File 2
                narration2 = str(header_row2.iloc[2]).strip()
                
                # Skip empty narrations
                if len(narration2) < 10 or narration2.lower() in ['nan', 'none', '']:
                    continue
                
                # Extract PO numbers from this narration
                po_matches2 = re.findall(PO_PATTERN, narration2)
                
                # Check if any of the lender POs are present in this borrower narration
                for po in po_matches1:
                    if po in po_matches2 and po not in found_pos:
                        # Found a matching PO, add this borrower transaction
                        matching_borrower_transactions.append({
                            'row': idx2,
                            'header_row': block_header2,
                            'amount': file2_credit,
                            'po': po,
                            'narration': narration2[:50] + '...'
                        })
                        total_borrower_amount += file2_credit
                        found_pos.add(po)
                        print(f"     Found matching PO '{po}' in File 2 Row {block_header2}")
                        print(f"      Borrower Amount: {file2_credit}, Narration: {narration2[:50]}...")
                        break
            
            # Check if we found ALL lender POs in borrower transactions
            if len(found_pos) == len(po_matches1):
                print(f"   All {len(po_matches1)} POs found in borrower transactions!")
                
                # Validate: Lender amount == Total borrower amount (exact match, no tolerance)
                if abs(file1_debit - total_borrower_amount) < 0.01:  # Allow small rounding differences
                    print(f"   Amount validation PASSED: Lender {file1_debit} == Total Borrower {total_borrower_amount}")
                    
                    # Generate next sequential Match ID using centralized manager
                    context = f"AggregatedPO_{po_matches1[0]}_Total_{len(matching_borrower_transactions)}_POs"
                    # Match ID will be assigned later in post-processing
                    match_id = None
                    
                    print(f"   AGGREGATED PO MATCH FOUND!")
                    
                    # Create the match
                    matches.append({
                        'match_id': match_id,
                        'Match_Type': 'Aggregated_PO',
                        'File1_Index': block_header1,
                        'File2_Index': [t['header_row'] for t in matching_borrower_transactions],
                        'PO_Count': len(po_matches1),
                        'All_POs': po_matches1,
                        'File1_Date': header_row1.iloc[0],
                        'File1_Description': header_row1.iloc[2],
                        'File1_Debit': file1_debit,
                        'File1_Credit': file1_credit,
                        'File2_Date': [t['row'] for t in matching_borrower_transactions],
                        'File2_Description': [t['narration'] for t in matching_borrower_transactions],
                        'File2_Debit': [0] * len(matching_borrower_transactions),  # Borrowers have 0 debit
                        'File2_Credit': [t['amount'] for t in matching_borrower_transactions],
                        'File1_Amount': file1_debit,
                        'File2_Amount': total_borrower_amount,
                        'File1_Type': 'Lender',
                        'File2_Type': 'Borrower',
                        'Lender_File': 1,
                        'Lender_Index': block_header1,
                        'Borrower_File': 2,
                        'Borrower_Index': [t['header_row'] for t in matching_borrower_transactions],
                        'Lender_Amount': file1_debit,
                        'Borrower_Amount': total_borrower_amount
                    })
                else:
                    print(f"   Amount validation FAILED: Lender {file1_debit} != Total Borrower {total_borrower_amount}")
            else:
                missing_pos = set(po_matches1) - found_pos
                print(f"   Missing POs in borrower transactions: {missing_pos}")
        
        print(f"\n=== AGGREGATED PO MATCHING RESULTS ===")
        print(f"Found {len(matches)} valid aggregated PO matches!")
        
        # Show some examples
        if matches:
            print("\n=== SAMPLE AGGREGATED PO MATCHES ===")
            for i, match in enumerate(matches[:3]):
                print(f"\nAggregated PO Match {i+1}:")
                print(f"Match ID: {match['match_id']}")
                print(f"PO Count: {match['PO_Count']}")
                print(f"POs: {', '.join(match['All_POs'][:5])}")
                print(f"Lender Amount: {match['Lender_Amount']}")
                print(f"Total Borrower Amount: {match['Borrower_Amount']}")
                print(f"File 1: {match['File1_Date']} - {str(match['File1_Description'])[:50]}...")
                print(f"  Type: {match['File1_Type']}, Debit: {match['File1_Debit']}, Credit: {match['File1_Credit']}")
        
        return matches
    
    # Transaction block identification methods are now provided by the shared TransactionBlockIdentifier instance
    # This ensures consistent behavior across all matching modules
