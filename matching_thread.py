"""
Matching Thread for Interunit Loan Matcher
Background processing logic for GUI application with sleep prevention
"""

import sys
import io
import time
from PySide6.QtCore import QThread, Signal
from interunit_loan_matcher import ExcelTransactionMatcher
from system_utils import SystemSleepPrevention, ProcessMonitor


class MatchingThread(QThread):
    """Background thread for running the matching process"""
    
    # Signals for communication with main thread
    progress_updated = Signal(int, str, int)  # step, status, matches_found
    step_completed = Signal(str, int)  # step_name, matches_found
    matching_finished = Signal(list, dict)  # matches, statistics
    error_occurred = Signal(str)  # error_message
    log_message = Signal(str)  # detailed log message
    
    def __init__(self, file1_path: str, file2_path: str):
        super().__init__()
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.is_cancelled = False
        self._original_stdout = None
        self._captured_output = None
        
        # Initialize system utilities
        self.sleep_prevention = SystemSleepPrevention()
        self.process_monitor = ProcessMonitor()
        
        # Store data for manual matching
        self.all_automatic_matches = None
        self.transactions1_data = None
        self.transactions2_data = None
        self.blocks1_data = None
        self.blocks2_data = None
    
    def _capture_print_output(self):
        """Start capturing print output"""
        self._original_stdout = sys.stdout
        self._captured_output = io.StringIO()
        sys.stdout = self._captured_output
    
    def _release_print_output(self):
        """Stop capturing and emit captured output"""
        if self._original_stdout and self._captured_output:
            sys.stdout = self._original_stdout
            captured_text = self._captured_output.getvalue()
            if captured_text.strip():
                # Split by lines and emit each line as a log message
                for line in captured_text.strip().split('\n'):
                    if line.strip():
                        self.log_message.emit(line.strip())
            self._captured_output.close()
            self._captured_output = None
            self._original_stdout = None
    
    def _start_system_protection(self):
        """Start system sleep prevention and monitoring"""
        try:
            self.sleep_prevention.start_preventing_sleep("Excel Transaction Matching")
            self.process_monitor.start_monitoring()
            self.log_message.emit("[PROTECTION] System sleep prevention activated")
            self.log_message.emit("[STATS] Process monitoring started")
        except Exception as e:
            self.log_message.emit(f"[WARNING] Warning: Could not start system protection: {e}")
    
    def _stop_system_protection(self):
        """Stop system sleep prevention and monitoring"""
        try:
            self.sleep_prevention.stop_preventing_sleep()
            self.process_monitor.stop_monitoring()
            self.log_message.emit("[PROTECTION] System sleep prevention deactivated")
        except Exception as e:
            self.log_message.emit(f"[WARNING] Warning: Could not stop system protection: {e}")
    
    def _update_activity(self):
        """Update process activity timestamp"""
        self.process_monitor.update_activity()
        
    def run(self):
        """Run the matching process in background thread with sleep prevention"""
        try:
            # Start system protection
            self._start_system_protection()
            
            # Start capturing print output
            self._capture_print_output()
            
            # PHASE 1: INITIALIZATION (0-5%)
            self.progress_updated.emit(2, "Initializing matcher...", 0)
            self.log_message.emit("START: Starting Interunit Loan Matcher...")
            self.log_message.emit(f"[FILE] File 1: {self.file1_path}")
            self.log_message.emit(f"[FILE] File 2: {self.file2_path}")
            
            # Reset unmatched tracker for new matching run
            try:
                from unmatched_tracker import reset_unmatched_tracker
                reset_unmatched_tracker()
                self.log_message.emit("   [OK] Reset unmatched record tracker")
            except ImportError:
                pass  # unmatched_tracker not available, continue anyway
            
            matcher = ExcelTransactionMatcher(self.file1_path, self.file2_path)
            self._update_activity()
            
            # Release any captured output so far
            self._release_print_output()
            
            if self.is_cancelled:
                return
                
            # PHASE 2: FILE PROCESSING (5-25%) - This is the heaviest part
            self.progress_updated.emit(8, "Loading Excel files...", 0)
            self.log_message.emit("[STATS] Processing Excel files...")
            self._update_activity()
            
            # Capture print output during file processing
            self._capture_print_output()
            self.log_message.emit("   - Loading Excel files and extracting data...")
            transactions1, transactions2, blocks1, blocks2, lc_numbers1, lc_numbers2, po_numbers1, po_numbers2, interunit_accounts1, interunit_accounts2, usd_amounts1, usd_amounts2 = matcher.process_files()
            self._release_print_output()
            self.log_message.emit(f"   [OK] File processing completed - {len(transactions1)} and {len(transactions2)} transactions loaded")
            self._update_activity()
            
            if self.is_cancelled:
                return
            
            # PHASE 2.5: GLOBAL PRE-FILTERING BY AMOUNT (OPTIMIZATION)
            self.progress_updated.emit(20, "Pre-filtering by amount...", 0)
            self.log_message.emit("[OPTIMIZATION] Global Pre-filtering by Amount")
            self.log_message.emit("   - Filtering transactions with matching amounts...")
            self.log_message.emit("   - This reduces dataset size for all matching modules...")
            self._update_activity()
            
            # Capture print output during pre-filtering
            self._capture_print_output()
            filtered_indices1, filtered_indices2, filtered_blocks1, filtered_blocks2 = matcher._prefilter_by_amount_matching(
                transactions1, transactions2, blocks1, blocks2
            )
            self._release_print_output()
            
            # Calculate reduction
            reduction1 = len(transactions1) - len(filtered_indices1)
            reduction2 = len(transactions2) - len(filtered_indices2)
            reduction_pct1 = (reduction1 / len(transactions1) * 100) if len(transactions1) > 0 else 0
            reduction_pct2 = (reduction2 / len(transactions2) * 100) if len(transactions2) > 0 else 0
            
            self.log_message.emit(f"   [OK] Pre-filtering complete:")
            self.log_message.emit(f"   - File 1: {len(filtered_indices1)}/{len(transactions1)} transactions ({reduction_pct1:.1f}% reduction)")
            self.log_message.emit(f"   - File 2: {len(filtered_indices2)}/{len(transactions2)} transactions ({reduction_pct2:.1f}% reduction)")
            self.log_message.emit(f"   - Blocks: {len(filtered_blocks1)}/{len(blocks1)} File 1, {len(filtered_blocks2)}/{len(blocks2)} File 2")
            self._update_activity()
            
            if self.is_cancelled:
                return
            
            # Filter data series to match filtered transactions (maintain original indices)
            # Use boolean indexing to filter while preserving original index positions
            filtered_lc_numbers1 = lc_numbers1.copy()
            filtered_lc_numbers2 = lc_numbers2.copy()
            filtered_po_numbers1 = po_numbers1.copy()
            filtered_po_numbers2 = po_numbers2.copy()
            filtered_interunit_accounts1 = interunit_accounts1.copy()
            filtered_interunit_accounts2 = interunit_accounts2.copy()
            filtered_usd_amounts1 = usd_amounts1.copy()
            filtered_usd_amounts2 = usd_amounts2.copy()
            
            # Set non-filtered indices to None (they won't be processed)
            for idx in range(len(lc_numbers1)):
                if idx not in filtered_indices1:
                    filtered_lc_numbers1.iloc[idx] = None
                    filtered_po_numbers1.iloc[idx] = None
                    filtered_interunit_accounts1.iloc[idx] = None
                    filtered_usd_amounts1.iloc[idx] = None
            
            for idx in range(len(lc_numbers2)):
                if idx not in filtered_indices2:
                    filtered_lc_numbers2.iloc[idx] = None
                    filtered_po_numbers2.iloc[idx] = None
                    filtered_interunit_accounts2.iloc[idx] = None
                    filtered_usd_amounts2.iloc[idx] = None
            
            # Update variables to use filtered data
            blocks1 = filtered_blocks1
            blocks2 = filtered_blocks2
            lc_numbers1 = filtered_lc_numbers1
            lc_numbers2 = filtered_lc_numbers2
            po_numbers1 = filtered_po_numbers1
            po_numbers2 = filtered_po_numbers2
            interunit_accounts1 = filtered_interunit_accounts1
            interunit_accounts2 = filtered_interunit_accounts2
            usd_amounts1 = filtered_usd_amounts1
            usd_amounts2 = filtered_usd_amounts2
            
            # Note: transactions1 and transactions2 remain unchanged (matching modules use original indices)
            # The filtered data series will naturally skip non-matching transactions
            
            # PHASE 3: MATCHING LOGIC (25-60%) - 5 steps, 7% each
            # Step 1: Narration Matching (25-32%)
            self.progress_updated.emit(25, "Finding narration matches...", 0)
            self.log_message.emit("[SEARCH] Step 1/5: Narration Matching")
            self.log_message.emit("   - Searching for exact text matches in transaction descriptions...")
            self.log_message.emit("   - Analyzing transaction descriptions for matches...")
            self._update_activity()
            
            narration_matches = matcher.narration_matching_logic.find_potential_matches(
                transactions1, transactions2, self.file1_path, self.file2_path, {}, None
            )
            self.log_message.emit(f"   [OK] Found {len(narration_matches)} narration matches")
            self.log_message.emit(f"   - Narration matching completed in {len(narration_matches)} matches")
            self.step_completed.emit("Narration Matching", len(narration_matches))
            
            self._update_activity()
            
            if self.is_cancelled:
                return
                
            # Step 2: LC Matching (32-39%)
            self.progress_updated.emit(32, "Finding LC matches...", 0)
            self.log_message.emit("[SEARCH] Step 2/5: LC Matching")
            self.log_message.emit("   - Filtering out already matched records...")
            
            # Create masks for unmatched records (after Narration matching)
            narration_matched_indices1 = set()
            narration_matched_indices2 = set()
            
            for match in narration_matches:
                narration_matched_indices1.add(match['File1_Index'])
                narration_matched_indices2.add(match['File2_Index'])
            
            # Filter LC numbers to only unmatched records
            lc_numbers1_unmatched = lc_numbers1.copy()
            lc_numbers2_unmatched = lc_numbers2.copy()
            
            for idx in narration_matched_indices1:
                if idx < len(lc_numbers1_unmatched):
                    lc_numbers1_unmatched.iloc[idx] = None
            
            for idx in narration_matched_indices2:
                if idx < len(lc_numbers2_unmatched):
                    lc_numbers2_unmatched.iloc[idx] = None
            
            self.log_message.emit("   - Searching for LC number matches...")
            self.log_message.emit("   - Analyzing LC numbers for potential matches...")
            lc_matches = matcher.lc_matching_logic.find_potential_matches(
                transactions1, transactions2, lc_numbers1_unmatched, lc_numbers2_unmatched,
                self.file1_path, self.file2_path, {}, None
            )
            self.log_message.emit(f"   [OK] Found {len(lc_matches)} LC matches")
            self.log_message.emit(f"   - LC matching completed in {len(lc_matches)} matches")
            self.step_completed.emit("LC Matching", len(lc_matches))
            
            if self.is_cancelled:
                return
                
            # Step 3: PO Matching (39-46%)
            self.progress_updated.emit(39, "Finding PO matches...", 0)
            self.log_message.emit("[SEARCH] Step 3/5: PO Matching")
            self.log_message.emit("   - Filtering out already matched records...")
            
            # Create masks for unmatched records (after Narration and LC matching)
            narration_lc_matched_indices1 = set()
            narration_lc_matched_indices2 = set()
            
            for match in narration_matches + lc_matches:
                narration_lc_matched_indices1.add(match['File1_Index'])
                narration_lc_matched_indices2.add(match['File2_Index'])
            
            # Filter PO numbers to only unmatched records
            po_numbers1_unmatched = po_numbers1.copy()
            po_numbers2_unmatched = po_numbers2.copy()
            
            for idx in narration_lc_matched_indices1:
                if idx < len(po_numbers1_unmatched):
                    po_numbers1_unmatched.iloc[idx] = None
            
            for idx in narration_lc_matched_indices2:
                if idx < len(po_numbers2_unmatched):
                    po_numbers2_unmatched.iloc[idx] = None
            
            self.log_message.emit("   - Searching for PO number matches...")
            self.log_message.emit("   - Analyzing PO numbers for potential matches...")
            po_matches = matcher.po_matching_logic.find_potential_matches(
                transactions1, transactions2, po_numbers1_unmatched, po_numbers2_unmatched,
                self.file1_path, self.file2_path, {}, None
            )
            self.log_message.emit(f"   [OK] Found {len(po_matches)} PO matches")
            self.log_message.emit(f"   - PO matching completed in {len(po_matches)} matches")
            self.step_completed.emit("PO Matching", len(po_matches))
            
            if self.is_cancelled:
                return
                
            # Step 4: Interunit Matching (46-53%)
            self.progress_updated.emit(46, "Finding interunit matches...", 0)
            self.log_message.emit("[SEARCH] Step 4/5: Interunit Matching")
            self.log_message.emit("   - Filtering out already matched records...")
            
            # Create masks for unmatched records (after Narration, LC, and PO matching)
            narration_lc_po_matched_indices1 = set()
            narration_lc_po_matched_indices2 = set()
            
            for match in narration_matches + lc_matches + po_matches:
                narration_lc_po_matched_indices1.add(match['File1_Index'])
                narration_lc_po_matched_indices2.add(match['File2_Index'])
            
            # Filter interunit accounts to only unmatched records
            interunit_accounts1_unmatched = interunit_accounts1.copy()
            interunit_accounts2_unmatched = interunit_accounts2.copy()
            
            for idx in narration_lc_po_matched_indices1:
                if idx < len(interunit_accounts1_unmatched):
                    interunit_accounts1_unmatched.iloc[idx] = None
            
            for idx in narration_lc_po_matched_indices2:
                if idx < len(interunit_accounts2_unmatched):
                    interunit_accounts2_unmatched.iloc[idx] = None
            
            self.log_message.emit("   - Searching for interunit account matches...")
            self.log_message.emit("   - Analyzing interunit accounts for potential matches...")
            interunit_matches = matcher.interunit_loan_matcher.find_potential_matches(
                transactions1, transactions2, interunit_accounts1_unmatched, interunit_accounts2_unmatched,
                self.file1_path, self.file2_path, {}, None
            )
            self.log_message.emit(f"   [OK] Found {len(interunit_matches)} interunit matches")
            self.log_message.emit(f"   - Interunit matching completed in {len(interunit_matches)} matches")
            self.step_completed.emit("Interunit Matching", len(interunit_matches))
            
            if self.is_cancelled:
                return
                
            # Step 5: Settlement Matching (53-60%)
            self.progress_updated.emit(53, "Finding settlement matches...", 0)
            self.log_message.emit("[SEARCH] Step 5/6: Settlement Matching")
            self.log_message.emit("   - Filtering out already matched records...")
            
            # Create masks for unmatched records (after Narration, LC, PO, and Interunit matching)
            narration_lc_po_interunit_matched_indices1 = set()
            narration_lc_po_interunit_matched_indices2 = set()
            
            for match in narration_matches + lc_matches + po_matches + interunit_matches:
                narration_lc_po_interunit_matched_indices1.add(match['File1_Index'])
                narration_lc_po_interunit_matched_indices2.add(match['File2_Index'])
            
            # Filter blocks to only unmatched records
            blocks1_unmatched = [b for b in blocks1 if b[0] not in narration_lc_po_interunit_matched_indices1]
            blocks2_unmatched = [b for b in blocks2 if b[0] not in narration_lc_po_interunit_matched_indices2]
            
            self.log_message.emit("   - Searching for settlement matches (Employee IDs)...")
            self.log_message.emit("   - Analyzing transaction narrations for Employee IDs...")
            settlement_matches = matcher.settlement_matching_logic.find_potential_matches(
                transactions1, transactions2, blocks1_unmatched, blocks2_unmatched,
                self.file1_path, self.file2_path
            )
            self.log_message.emit(f"   [OK] Found {len(settlement_matches)} settlement matches")
            self.log_message.emit(f"   - Settlement matching completed in {len(settlement_matches)} matches")
            self.step_completed.emit("Settlement Matching", len(settlement_matches))
            
            if self.is_cancelled:
                return
                
            # Step 6: USD Matching (60-67%)
            self.progress_updated.emit(60, "Finding USD matches...", 0)
            self.log_message.emit("[SEARCH] Step 6/6: USD Matching")
            self.log_message.emit("   - Filtering out already matched records...")
            
            # Create masks for unmatched records (after Narration, LC, PO, Interunit, and Settlement matching)
            all_matched_indices1 = set()
            all_matched_indices2 = set()
            
            for match in narration_matches + lc_matches + po_matches + interunit_matches + settlement_matches:
                all_matched_indices1.add(match['File1_Index'])
                all_matched_indices2.add(match['File2_Index'])
            
            # Filter USD amounts to only unmatched records
            usd_amounts1_unmatched = usd_amounts1.copy()
            usd_amounts2_unmatched = usd_amounts2.copy()
            
            for idx in all_matched_indices1:
                if idx < len(usd_amounts1_unmatched):
                    usd_amounts1_unmatched.iloc[idx] = None
            
            for idx in all_matched_indices2:
                if idx < len(usd_amounts2_unmatched):
                    usd_amounts2_unmatched.iloc[idx] = None
            
            self.log_message.emit("   - Searching for USD amount matches...")
            self.log_message.emit("   - Analyzing USD amounts for potential matches...")
            usd_matches = matcher.usd_matching_logic.find_potential_matches(
                transactions1, transactions2, usd_amounts1_unmatched, usd_amounts2_unmatched,
                self.file1_path, self.file2_path, {}, None
            )
            self.log_message.emit(f"   [OK] Found {len(usd_matches)} USD matches")
            self.log_message.emit(f"   - USD matching completed in {len(usd_matches)} matches")
            self.step_completed.emit("USD Matching", len(usd_matches))
            
            if self.is_cancelled:
                return
                
            # PHASE 4: MATCH PROCESSING (67-75%)
            self.progress_updated.emit(67, "Processing matches...", 0)
            self.log_message.emit("[STATS] Processing all matches...")
            
            # Combine all matches
            self.log_message.emit("   - Combining matches from all matching types...")
            all_matches = narration_matches + lc_matches + po_matches + interunit_matches + settlement_matches + usd_matches
            self.log_message.emit(f"   - Combined {len(all_matches)} total matches")
            
            # Assign sequential Match IDs
            self.log_message.emit("   - Assigning sequential Match IDs...")
            self.log_message.emit("   - Organizing matches for output...")
            match_counter = 1
            for match in all_matches:
                match_id = f"M{match_counter:03d}"
                match['match_id'] = match_id
                match_counter += 1
            
            # Sort matches by the newly assigned sequential Match IDs
            self.log_message.emit("   - Sorting matches by Match ID...")
            all_matches.sort(key=lambda x: x['match_id'])
            self.log_message.emit("   - Match processing completed successfully!")
            
            # Create statistics
            stats = {
                'total_matches': len(all_matches),
                'narration_matches': len(narration_matches),
                'lc_matches': len(lc_matches),
                'po_matches': len(po_matches),
                'interunit_matches': len(interunit_matches),
                'settlement_matches': len(settlement_matches),
                'usd_matches': len(usd_matches),
                'manual_matches': 0  # Manual matches added later
            }
            
            self.log_message.emit("[DONE] Matching completed successfully!")
            self.log_message.emit(f"[RESULTS] Final Results: {stats['total_matches']} total matches found")
            self.progress_updated.emit(75, "Matching completed successfully!", stats['total_matches'])
            
            # Store matches and data for potential manual matching
            self.all_automatic_matches = all_matches
            self.transactions1_data = transactions1
            self.transactions2_data = transactions2
            self.blocks1_data = blocks1
            self.blocks2_data = blocks2
            
            self.matching_finished.emit(all_matches, stats)
            
        except Exception as e:
            # Make sure to release print output capture on error
            self._release_print_output()
            self.log_message.emit(f"[ERROR] Error occurred: {str(e)}")
            self.error_occurred.emit(str(e))
            
        finally:
            # Always stop system protection
            self._stop_system_protection()
    
    def cancel(self):
        """Cancel the matching process"""
        self.is_cancelled = True
