import pandas as pd
import re

# USD Amount extraction pattern
# Matches USD amounts in various formats:
# - $.789,663.20 (starts with $., comma, decimal)
# - $6,400 (starts with $, comma, no decimal)
# - $339520.78 (starts with $, no comma, decimal)
# - $80 (starts with $, no comma, no decimal)
# - $147,401.28 (standard format)
# - $80 (simple format)
USD_PATTERN = r'\$\s*\.?\s*[\d,]+\.?\d*'

class USDMatchingLogic:
    """Handles the logic for finding USD amount matches between two files."""
    
    def __init__(self):
        pass
    
    def find_potential_matches(self, transactions1, transactions2, usd_amounts1, usd_amounts2, existing_matches=None, match_counter=0):
        """Find potential USD amount matches between the two files."""
        # Filter rows with USD amounts
        usd_transactions1 = transactions1[usd_amounts1.notna()].copy()
        usd_transactions2 = transactions2[usd_amounts2.notna()].copy()
        
        print(f"\nFile 1: {len(usd_transactions1)} transactions with USD amounts")
        print(f"File 2: {len(usd_transactions2)} transactions with USD amounts")
        
        # Find matches - USD Amount → Transaction Amount matching
        matches = []
        
        # Use shared state if provided, otherwise create new
        if existing_matches is None:
            existing_matches = {}
        if match_counter is None:
            match_counter = 0
        
        print(f"\n=== USD MATCHING LOGIC ===")
        print(f"1. Check if lender debit and borrower credit amounts are EXACTLY the same")
        print(f"2. Check if BOTH narrations have the SAME NUMBER of USD amounts")
        print(f"3. Check if ALL USD amounts are IDENTICAL between lender and borrower")
        print(f"4. Only if all criteria match: Assign same Match ID")
        print(f"5. IMPORTANT: All transactions with same USD+Amount get SAME Match ID")
        
        # Use shared state for tracking which combinations have already been matched
        # Key: (USD_Amount, Transaction_Amount), Value: match_id
        
        # Process each transaction in File 1 to find matches in File 2
        for idx1, usd1 in enumerate(usd_amounts1):
            if not usd1:
                continue
                
            print(f"\n--- Processing File 1 Row {idx1} with USD: {usd1} ---")
            
            # Find the transaction block header row for this USD in File 1
            block_header1 = self.find_transaction_block_header(idx1, transactions1)
            header_row1 = transactions1.iloc[block_header1]
            
            # Extract amounts and determine transaction type for File 1
            # Based on investigation: amounts are in columns 8 and 9 (iloc[7] and iloc[8])
            file1_debit = header_row1.iloc[7] if pd.notna(header_row1.iloc[7]) else 0
            file1_credit = header_row1.iloc[8] if pd.notna(header_row1.iloc[8]) else 0
            
            file1_is_lender = file1_debit > 0
            file1_is_borrower = file1_credit > 0
            file1_amount = file1_debit if file1_is_lender else file1_credit
            
            print(f"  File 1: Amount={file1_amount}, Type={'Lender' if file1_is_lender else 'Borrower'}")
            
            # Now look for matches in File 2
            for idx2, usd2 in enumerate(usd_amounts2):
                if not usd2:
                    continue
                    
                print(f"    Checking File 2 Row {idx2} with USD: {usd2}")
                
                # Find the transaction block header row for this USD in File 2
                block_header2 = self.find_transaction_block_header(idx2, transactions2)
                header_row2 = transactions2.iloc[block_header2]
                
                # Extract amounts and determine transaction type for File 2
                # Based on investigation: amounts are in columns 8 and 9 (iloc[7] and iloc[8])
                file2_debit = header_row2.iloc[7] if pd.notna(header_row2.iloc[7]) else 0
                file2_credit = header_row2.iloc[8] if pd.notna(header_row2.iloc[8]) else 0
                
                file2_is_lender = file2_debit > 0
                file2_is_borrower = file2_credit > 0
                file2_amount = file2_debit if file2_is_lender else file2_credit
                
                print(f"      File 2: Amount={file2_amount}, Type={'Lender' if file2_is_lender else 'Borrower'}")
                
                # STEP 1: Check if amounts are EXACTLY the same
                if file1_amount != file2_amount:
                    print(f"      ❌ REJECTED: Amounts don't match ({file1_amount} vs {file2_amount})")
                    continue
                
                print(f"      ✅ STEP 1 PASSED: Amounts match exactly")
                
                # STEP 2: Check if transaction types are opposite (one lender, one borrower)
                if not ((file1_is_lender and file2_is_borrower) or (file1_is_borrower and file2_is_lender)):
                    print(f"      ❌ REJECTED: Transaction types don't match (both same type)")
                    continue
                
                print(f"      ✅ STEP 2 PASSED: Transaction types are opposite")
                
                # STEP 3: Check if USD amounts match
                if usd1 != usd2:
                    print(f"      ❌ REJECTED: USD amounts don't match ('{usd1}' vs '{usd2}')")
                    continue
                
                print(f"      ✅ STEP 3 PASSED: USD amounts match")
                
                # STEP 4: Check if both narrations have the same number of USD amounts
                # Extract all USD amounts from both narrations
                narration1 = str(header_row1.iloc[2]).upper()
                narration2 = str(header_row2.iloc[2]).upper()
                
                # DEBUG: Show what we're trying to match
                print(f"      DEBUG: File 1 narration: {narration1[:100]}...")
                print(f"      DEBUG: File 2 narration: {narration2[:100]}...")
                print(f"      DEBUG: Using USD_PATTERN: {USD_PATTERN}")
                
                usd_amounts_in_narration1 = re.findall(USD_PATTERN, narration1)
                usd_amounts_in_narration2 = re.findall(USD_PATTERN, narration2)
                
                print(f"      File 1 narration has {len(usd_amounts_in_narration1)} USD amounts: {usd_amounts_in_narration1}")
                print(f"      File 2 narration has {len(usd_amounts_in_narration2)} USD amounts: {usd_amounts_in_narration2}")
                
                # FIX: If regex extraction fails, use the actual USD amounts that triggered the match
                if not usd_amounts_in_narration1:
                    print(f"      ⚠️  WARNING: Regex didn't find USD amounts in File 1 narration, using actual USD amount: {usd1}")
                    usd_amounts_in_narration1 = [usd1]
                
                if not usd_amounts_in_narration2:
                    print(f"      ⚠️  WARNING: Regex didn't find USD amounts in File 2 narration, using actual USD amount: {usd2}")
                    usd_amounts_in_narration2 = [usd2]
                
                if len(usd_amounts_in_narration1) != len(usd_amounts_in_narration2):
                    print(f"      ❌ REJECTED: Different number of USD amounts ({len(usd_amounts_in_narration1)} vs {len(usd_amounts_in_narration2)})")
                    continue
                
                print(f"      ✅ STEP 4 PASSED: Same number of USD amounts")
                
                # STEP 5: Check if ALL USD amounts are identical between narrations
                # Sort both lists to ensure order doesn't matter
                sorted_usd1 = sorted(usd_amounts_in_narration1)
                sorted_usd2 = sorted(usd_amounts_in_narration2)
                
                if sorted_usd1 != sorted_usd2:
                    print(f"      ❌ REJECTED: USD amounts don't match exactly")
                    print(f"        File 1: {sorted_usd1}")
                    print(f"        File 2: {sorted_usd2}")
                    continue
                
                print(f"      ✅ STEP 5 PASSED: All USD amounts are identical")
                
                # STEP 6: Check if we already have a match for this combination
                match_key = (usd1, file1_amount)
                
                if match_key in existing_matches:
                    # Use existing Match ID for consistency
                    match_id = existing_matches[match_key]
                    print(f"      🔄 REUSING existing Match ID: {match_id}")
                else:
                    # Create new Match ID
                    match_counter += 1
                    match_id = f"M{match_counter:03d}"
                    existing_matches[match_key] = match_id
                    print(f"      🆕 CREATING new Match ID: {match_id}")
                
                print(f"      🎉 ALL CRITERIA MET - USD MATCH FOUND!")
                
                # Create the match
                matches.append({
                    'match_id': match_id,
                    'Match_Type': 'USD',  # Add explicit match type
                    'File1_Index': block_header1,
                    'File2_Index': block_header2,
                    'USD_Amount': usd1,
                    'File1_Date': header_row1.iloc[0],
                    'File1_Description': header_row1.iloc[2],
                    'File1_Debit': header_row1.iloc[7],
                    'File1_Credit': header_row1.iloc[8],
                    'File2_Date': header_row2.iloc[0],
                    'File2_Description': header_row2.iloc[2],
                    'File2_Debit': header_row2.iloc[7],
                    'File2_Credit': header_row2.iloc[8],
                    'File1_Amount': file1_amount,
                    'File2_Amount': file2_amount,
                    'File1_Type': 'Lender' if file1_is_lender else 'Borrower',
                    'File2_Type': 'Lender' if file2_is_lender else 'Borrower',
                    'USD_Count': len(usd_amounts_in_narration1),
                    'USD_Amounts_File1': usd_amounts_in_narration1,
                    'USD_Amounts_File2': usd_amounts_in_narration2
                })
        
        print(f"\n=== USD MATCHING RESULTS ===")
        print(f"Found {len(matches)} valid USD matches across {len(existing_matches)} unique Match ID combinations!")
        
        # Show some examples
        if matches:
            print("\n=== SAMPLE USD MATCHES ===")
            for i, match in enumerate(matches[:5]):
                print(f"\nUSD Match {i+1}:")
                print(f"Match ID: {match['match_id']}")
                print(f"USD Amount: {match['USD_Amount']}")
                print(f"Transaction Amount: {match['File1_Amount']}")
                print(f"USD Count: {match['USD_Count']}")
                print(f"File 1: {match['File1_Date']} - {str(match['File1_Description'])[:50]}...")
                print(f"  Type: {match['File1_Type']}, Debit: {match['File1_Debit']}, Credit: {match['File1_Credit']}")
                print(f"  USD Amounts: {match['USD_Amounts_File1']}")
                print(f"File 2: {match['File2_Date']} - {str(match['File2_Description'])[:50]}...")
                print(f"  Type: {match['File2_Type']}, Debit: {match['File2_Debit']}, Credit: {match['File2_Credit']}")
                print(f"  USD Amounts: {match['USD_Amounts_File2']}")
        
        return matches
    
    def find_transaction_block_header(self, description_row_idx, transactions_df):
        """Find the transaction block header row for a given description row."""
        # Start from the description row and go backwards to find the block header
        # Block header is the row with date and particulars (Dr/Cr)
        for row_idx in range(description_row_idx, -1, -1):
            row = transactions_df.iloc[row_idx]
            
            # Check if this row has a date and particulars
            has_date = pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() != ''
            has_particulars = pd.notna(row.iloc[1]) and str(row.iloc[1]).strip() != ''
            
            # Check if this row has either Debit or Credit amount (not both nan)
            # Based on investigation: amounts are in columns 8 and 9 (iloc[7] and iloc[8])
            has_debit = pd.notna(row.iloc[7]) and row.iloc[7] != 0
            has_credit = pd.notna(row.iloc[8]) and row.iloc[8] != 0
            
            # Transaction block header: has date, particulars, and either debit or credit
            if has_date and (has_debit or has_credit):
                return row_idx
        
        # If no header found, return the description row itself
        return description_row_idx
