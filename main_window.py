"""
Main Window for Interunit Loan Matcher
Central window class for the GUI application
"""

import os
from pathlib import Path
from typing import Optional
from datetime import datetime

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QMessageBox, QTabWidget, QFrame, QStackedWidget
)
from PySide6.QtCore import Qt

from gui_widgets import FileSelectionWidget, ProcessingWidget, ResultsWidget, LogWidget
from matching_thread import MatchingThread
# No custom styling - using macOS native style
from interunit_loan_matcher import ExcelTransactionMatcher
from bank_account_mapping_widget import BankAccountMappingWidget
from manual_matching_window import ManualMatchingWidget
from matching_logic import ManualMatchingLogic
from transaction_block_identifier import TransactionBlockIdentifier


class MainWindow(QMainWindow):
    """Main application window"""
    
    def __init__(self):
        super().__init__()
        self.matching_thread = None
        self.current_file1 = ""
        self.current_file2 = ""
        self.current_matches = []
        self.start_time = None
        self.files_saved = False  # Track if output files have been saved
        
        self.init_ui()
        self.apply_styling()
        self._preload_test_files()
        
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("Interunit Loan Matcher - GUI")
        self.setGeometry(100, 100, 1200, 650)
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Create main layout
        main_layout = QVBoxLayout()
        main_layout.setSpacing(5)
        main_layout.setContentsMargins(0, 0, 0, 0)  # No padding - sections have universal 15px padding
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        
        # Tab 1: Matching
        matching_tab = QWidget()
        matching_layout = QVBoxLayout()
        matching_layout.setSpacing(5)
        matching_layout.setContentsMargins(0, 0, 0, 0)
        
        # Top row: File selection and Match steps side by side
        top_row = QHBoxLayout()
        top_row.setSpacing(0)  # No spacing - vertical line will provide separation
        top_row.setContentsMargins(0, 0, 0, 0)  # No margins for consistent alignment
        top_row.setAlignment(Qt.AlignmentFlag.AlignTop)  # Align widgets at the top
        
        # File selection widget (left) - equal width
        self.file_selection = FileSelectionWidget()
        self.file_selection.files_selected.connect(self.on_files_selected)
        self.file_selection.run_match_button.clicked.connect(self.start_matching)
        top_row.addWidget(self.file_selection, 1, Qt.AlignmentFlag.AlignTop)  # Align at top
        
        # Vertical line separator between Select section and Progress section
        vertical_separator = QFrame()
        vertical_separator.setFrameShape(QFrame.Shape.VLine)
        vertical_separator.setFrameShadow(QFrame.Shadow.Sunken)
        vertical_separator.setStyleSheet("color: #CCCCCC; background-color: #CCCCCC;")
        vertical_separator.setLineWidth(2)
        vertical_separator.setFixedWidth(2)
        top_row.addWidget(vertical_separator)
        
        # Processing widget (right) - equal width
        self.processing_widget = ProcessingWidget()
        top_row.addWidget(self.processing_widget, 1, Qt.AlignmentFlag.AlignTop)  # Align at top
        
        # Get reference to overall progress bar from processing widget
        self.overall_progress = self.processing_widget.overall_progress
        
        matching_layout.addLayout(top_row)
        
        # Horizontal line separator between top row and Match Summary
        horizontal_separator = QFrame()
        horizontal_separator.setFrameShape(QFrame.Shape.HLine)
        horizontal_separator.setFrameShadow(QFrame.Shadow.Sunken)
        horizontal_separator.setStyleSheet("color: #CCCCCC; background-color: #CCCCCC;")
        horizontal_separator.setLineWidth(2)
        horizontal_separator.setFixedHeight(2)
        matching_layout.addWidget(horizontal_separator)
        
        # Match Summary section - full width
        matching_layout.setSpacing(5)  # Reset spacing to normal after separator
        self.results_widget = ResultsWidget()
        # Connect the open buttons to the main window's methods
        self.results_widget.open_folder_button.clicked.connect(self.open_output_folder)
        self.results_widget.open_files_button.clicked.connect(self.open_output_files)
        matching_layout.addWidget(self.results_widget)
        
        # Horizontal line separator between Match Summary and Process Log
        horizontal_separator2 = QFrame()
        horizontal_separator2.setFrameShape(QFrame.Shape.HLine)
        horizontal_separator2.setFrameShadow(QFrame.Shadow.Sunken)
        horizontal_separator2.setStyleSheet("color: #CCCCCC; background-color: #CCCCCC;")
        horizontal_separator2.setLineWidth(2)
        horizontal_separator2.setFixedHeight(2)
        matching_layout.addWidget(horizontal_separator2)
        
        # Bottom stack for Process Log and Manual Matching
        self.bottom_stack = QStackedWidget()
        
        # Process Log section - full width
        self.log_widget = LogWidget()
        self.bottom_stack.addWidget(self.log_widget)  # Index 0
        
        # Manual Matching section - full width
        self.manual_matching_widget = ManualMatchingWidget()
        self.manual_matching_widget.finished.connect(self.on_manual_matches_finished)
        self.manual_matching_widget.skipped.connect(self.on_manual_matches_skipped)
        self.manual_matching_widget.cancelled.connect(self.on_manual_matches_cancelled)
        self.manual_matching_widget.match_confirmed.connect(self.on_manual_match_confirmed)
        self.bottom_stack.addWidget(self.manual_matching_widget)  # Index 1
        
        matching_layout.addWidget(self.bottom_stack, 1) # Give bottom stack stretch factor 1 to maximize height
        
        matching_tab.setLayout(matching_layout)
        self.tab_widget.addTab(matching_tab, "Matching")
        
        # Tab 2: Bank Account Mapping
        self.bank_mapping_widget = BankAccountMappingWidget()
        self.bank_mapping_widget.mapping_changed.connect(self.on_mapping_changed)
        self.tab_widget.addTab(self.bank_mapping_widget, "Bank Account Mapping")
        
        main_layout.addWidget(self.tab_widget)
        central_widget.setLayout(main_layout)
        
        # Apply styling
        # No custom styling - using macOS native style
        
        # Add initial log message
        self.log_widget.add_log("Application started. Please select Excel files to begin.")
        self.log_widget.add_log("<b><font color='red'>💡 TIP: For long processing sessions, consider disabling PC sleep mode to prevent interruptions during matching.</font></b>")
    
    def on_mapping_changed(self):
        """Handle bank account mapping changes"""
        self.log_widget.add_log("📝 Bank account mappings updated. Changes will be applied to future matching operations.")
    
    def _preload_test_files(self):
        """Preload test files for development/testing"""
        try:
            # Get the directory where the script is located
            script_dir = Path(__file__).parent
            
            # Define test file paths
            file1_path = script_dir / "Resources" / "New Geo to Steel.xlsx"
            file2_path = script_dir / "Resources" / "New Steel to Geo.xlsx"
            
            # Check if both files exist
            if file1_path.exists() and file2_path.exists():
                # Set the files in the file selection widget
                self.file_selection.set_files(str(file1_path), str(file2_path))
                # Also set the current file paths
                self.current_file1 = str(file1_path)
                self.current_file2 = str(file2_path)
                self.file1_path = str(file1_path)
                self.file2_path = str(file2_path)
                
                self.log_widget.add_log(f"✅ Preloaded test files:")
                self.log_widget.add_log(f"   File 1: {file1_path.name}")
                self.log_widget.add_log(f"   File 2: {file2_path.name}")
            else:
                # Log which files are missing
                missing_files = []
                if not file1_path.exists():
                    missing_files.append(f"File 1: {file1_path.name}")
                if not file2_path.exists():
                    missing_files.append(f"File 2: {file2_path.name}")
                self.log_widget.add_log(f"⚠️ Test files not found. Missing: {', '.join(missing_files)}")
        except Exception as e:
            # Silently fail if preload doesn't work - don't interrupt normal operation
            pass
    
    def apply_styling(self):
        """Apply minimal styling for button bold text"""
        from gui_styles import get_main_stylesheet
        self.setStyleSheet(get_main_stylesheet())
    
    def on_files_selected(self, file1_path: str, file2_path: str):
        """Handle file selection"""
        self.current_file1 = file1_path
        self.current_file2 = file2_path
        # Also set the attributes that open_output_folder expects
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.log_widget.add_log(f"Files selected: {os.path.basename(file1_path)} and {os.path.basename(file2_path)}")
    
    def start_matching(self):
        """Start the matching process"""
        if not self.current_file1 or not self.current_file2:
            QMessageBox.warning(self, "No Files", "Please select both Excel files first.")
            return
        
        # Record start time
        self.start_time = datetime.now()
        start_time_str = self.start_time.strftime("%Y-%m-%d %H:%M:%S")
        
        self.log_widget.add_log("="*60)
        self.log_widget.add_log(f"🚀 MATCHING PROCESS STARTED")
        self.log_widget.add_log(f"⏰ Start Time: {start_time_str}")
        self.log_widget.add_log("="*60)
        self.log_widget.add_log("Starting matching process...")
        
        # Reset state for new matching process
        self.files_saved = False
        self.processing_widget.set_processing_state(True)
        self.processing_widget.reset_progress()
        self.results_widget.reset_results()
        self.overall_progress.setValue(0)
        
        # Create and start matching thread
        self.matching_thread = MatchingThread(self.current_file1, self.current_file2)
        self.matching_thread.progress_updated.connect(self.update_overall_progress)
        self.matching_thread.step_completed.connect(self.processing_widget.complete_step)
        self.matching_thread.matching_finished.connect(self.on_matching_finished)
        self.matching_thread.error_occurred.connect(self.on_matching_error)
        self.matching_thread.log_message.connect(self.log_widget.add_log)
        self.matching_thread.start()
    
    def update_overall_progress(self, step: int, status: str, matches_found: int):
        """Update overall progress bar"""
        self.overall_progress.setValue(step)
        self.log_widget.add_log(f"{status} ({matches_found} matches found)")
    
    def on_matching_finished(self, matches: list, statistics: dict):
        """Handle matching completion"""
        self.current_matches = matches
        self.files_saved = False  # Reset files saved state
        self.current_statistics = statistics  # Store statistics for real-time updates
        self.results_widget.update_results(statistics, enable_buttons=False)  # Don't enable buttons yet
        
        self.log_widget.add_log(f"Matching completed successfully! Found {statistics['total_matches']} matches.")
        
        # Check for potential manual matches (before generating output files)
        if matches and len(matches) > 0:
            # Get data from matching thread for manual matching
            if (self.matching_thread and 
                hasattr(self.matching_thread, 'transactions1_data') and 
                self.matching_thread.transactions1_data is not None):
                self.log_widget.add_log("🔍 Checking for potential manual matches...")
                self._check_manual_matches(matches)
            else:
                # Fallback: generate output files directly
                self.create_output_files_with_progress(matches)
        else:
            # Record finish time even when no matches found
            finish_time = datetime.now()
            finish_time_str = finish_time.strftime("%Y-%m-%d %H:%M:%S")
            
            if self.start_time:
                total_duration = finish_time - self.start_time
                hours, remainder = divmod(total_duration.total_seconds(), 3600)
                minutes, seconds = divmod(remainder, 60)
                
                if hours > 0:
                    duration_str = f"{int(hours)}h {int(minutes)}m {seconds:.1f}s"
                elif minutes > 0:
                    duration_str = f"{int(minutes)}m {seconds:.1f}s"
                else:
                    duration_str = f"{seconds:.1f}s"
            else:
                duration_str = "Unknown"
            
            self.log_widget.add_log("="*60)
            self.log_widget.add_log(f"🎉 MATCHING PROCESS COMPLETED")
            self.log_widget.add_log(f"⏰ Finish Time: {finish_time_str}")
            self.log_widget.add_log(f"⏱️  Total Duration: {duration_str}")
            self.log_widget.add_log("="*60)
            self.log_widget.add_log("No matches found. No output files created.")
            self.processing_widget.set_processing_state(False)
    
    def _check_manual_matches(self, automatic_matches: list):
        """Check for potential manual matches and switch to manual matching view if needed"""
        try:
            # Store automatic matches to combine later
            self.temp_automatic_matches = automatic_matches
            
            # Get data from matching thread
            transactions1 = self.matching_thread.transactions1_data
            transactions2 = self.matching_thread.transactions2_data
            blocks1 = self.matching_thread.blocks1_data
            blocks2 = self.matching_thread.blocks2_data
            
            # Initialize manual matching logic
            block_identifier = TransactionBlockIdentifier()
            manual_logic = ManualMatchingLogic(block_identifier)
            
            # Find potential manual matches
            potential_matches = manual_logic.find_potential_manual_matches(
                transactions1, transactions2, blocks1, blocks2,
                self.current_file1, self.current_file2, automatic_matches
            )
            
            if potential_matches and len(potential_matches) > 0:
                self.log_widget.add_log(f"📋 Found {len(potential_matches)} potential manual match pairs")
                self.log_widget.add_log("   Switching to manual matching review...")
                
                # Update progress bar to show manual matching phase
                self.overall_progress.setValue(76)
                
                # Load potential matches into the widget and switch view
                self.manual_matching_widget.load_potential_matches(
                    potential_matches, self.current_file1, self.current_file2
                )
                self.bottom_stack.setCurrentIndex(1)  # Show Manual Matching Widget
            else:
                # No potential manual matches, generate output files with automatic matches
                self.log_widget.add_log("   No potential manual matches found.")
                self.create_output_files_with_progress(automatic_matches)
                
        except Exception as e:
            self.log_widget.add_log(f"❌ Error during manual matching check: {str(e)}")
            # Fallback: generate output files with automatic matches
            self.create_output_files_with_progress(automatic_matches)

    def on_manual_match_confirmed(self, count: int):
        """Handle real-time manual match confirmation"""
        # Update manual match count in real-time
        # Calculate automatic matches total from current statistics
        if hasattr(self, 'current_statistics') and self.current_statistics:
            automatic_total = (
                self.current_statistics.get('narration_matches', 0) +
                self.current_statistics.get('lc_matches', 0) +
                self.current_statistics.get('po_matches', 0) +
                self.current_statistics.get('interunit_matches', 0) +
                self.current_statistics.get('settlement_matches', 0) +
                self.current_statistics.get('usd_matches', 0)
            )
            self.results_widget.update_manual_match_count(count, automatic_total)
        else:
            # Fallback: update without automatic total
            self.results_widget.update_manual_match_count(count)

    def on_manual_matches_finished(self, confirmed_manual_matches):
        """Handle completion of manual matching review"""
        self.log_widget.add_log(f"✅ Manual matching completed: {len(confirmed_manual_matches)} matches confirmed")
        
        # Combine automatic and manual matches
        all_matches = self.temp_automatic_matches + confirmed_manual_matches
        
        # Assign sequential Match IDs to all matches
        match_counter = 1
        for match in all_matches:
            match_id = f"M{match_counter:03d}"
            match['match_id'] = match_id
            match_counter += 1
        
        # Sort matches by Match ID
        all_matches.sort(key=lambda x: x['match_id'])
        
        # Update statistics to include manual matches
        # Count matches by type from all_matches
        narration_count = sum(1 for m in all_matches if m.get('Match_Type') == 'Narration')
        lc_count = sum(1 for m in all_matches if m.get('Match_Type') == 'LC')
        po_count = sum(1 for m in all_matches if m.get('Match_Type') == 'PO')
        interunit_count = sum(1 for m in all_matches if m.get('Match_Type') == 'Interunit')
        settlement_count = sum(1 for m in all_matches if m.get('Match_Type') == 'Settlement')
        usd_count = sum(1 for m in all_matches if m.get('Match_Type') == 'USD')
        manual_count = sum(1 for m in all_matches if m.get('Match_Type') == 'Manual')
        
        updated_statistics = {
            'total_matches': len(all_matches),
            'narration_matches': narration_count,
            'lc_matches': lc_count,
            'po_matches': po_count,
            'interunit_matches': interunit_count,
            'settlement_matches': settlement_count,
            'usd_matches': usd_count,
            'manual_matches': manual_count
        }
        
        # Update results display with new statistics
        self.results_widget.update_results(updated_statistics, enable_buttons=False)
        
        # Update progress bar to show manual matching complete
        self.overall_progress.setValue(80)
        
        # Switch back to log view
        self.bottom_stack.setCurrentIndex(0)
        
        # Generate output files with combined matches
        self.create_output_files_with_progress(all_matches)

    def on_manual_matches_skipped(self):
        """Handle user skipping manual matching"""
        self.log_widget.add_log("⏭️ Manual matching skipped. Generating output files with automatic matches only.")
        # Update progress bar
        self.overall_progress.setValue(80)
        self.bottom_stack.setCurrentIndex(0)  # Switch back to log view
        self.create_output_files_with_progress(self.temp_automatic_matches)

    def on_manual_matches_cancelled(self):
        """Handle user cancelling manual matching"""
        self.log_widget.add_log("⚠️ Manual matching cancelled. Generating output files with automatic matches only.")
        # Update progress bar
        self.overall_progress.setValue(80)
        self.bottom_stack.setCurrentIndex(0)  # Switch back to log view
        self.create_output_files_with_progress(self.temp_automatic_matches)
    
    def create_output_files_with_progress(self, matches):
        """Create output files with accurate progress tracking"""
        try:
            # Keep processing state active during file creation
            self.processing_widget.set_processing_state(True)
            
            # PHASE 1: PREPARATION (80-85%)
            self.overall_progress.setValue(80)
            self.log_widget.add_log("📁 Preparing file creation...")
            self.log_widget.add_log(f"   - Processing {len(matches)} matches for output files")
            
            from interunit_loan_matcher import ExcelTransactionMatcher
            matcher = ExcelTransactionMatcher(self.current_file1, self.current_file2)
            
            # PHASE 2: LOAD TRANSACTION DATA (85-90%) - This takes significant time
            self.overall_progress.setValue(85)
            self.log_widget.add_log("📖 Loading first transaction file...")
            self.log_widget.add_log(f"   - Reading: {self.current_file1}")
            self.log_widget.add_log("   - Processing Excel file structure...")
            matcher.metadata1, matcher.transactions1 = matcher.read_complex_excel(self.current_file1)
            self.log_widget.add_log(f"   ✅ Loaded {len(matcher.transactions1)} transactions from File 1")
            
            self.overall_progress.setValue(88)
            self.log_widget.add_log("📖 Loading second transaction file...")
            self.log_widget.add_log(f"   - Reading: {self.current_file2}")
            self.log_widget.add_log("   - Processing Excel file structure...")
            matcher.metadata2, matcher.transactions2 = matcher.read_complex_excel(self.current_file2)
            self.log_widget.add_log(f"   ✅ Loaded {len(matcher.transactions2)} transactions from File 2")
            
            # PHASE 3: CREATE MATCHED FILES (90-100%) - This is the longest part
            self.overall_progress.setValue(90)
            self.log_widget.add_log("📝 Creating matched Excel files...")
            self.log_widget.add_log("   - Generating output files with matched transactions...")
            self.log_widget.add_log("   - This may take 30+ seconds for large files...")
            self.log_widget.add_log("   - Creating Excel workbooks and formatting...")
            
            # This is where most of the time is spent - creating and formatting Excel files
            matcher.create_matched_files(matches, matcher.transactions1, matcher.transactions2)
            self.log_widget.add_log("   - Excel file generation completed!")
            
            # PHASE 4: COMPLETE (100%)
            self.overall_progress.setValue(100)
            
            # Record finish time and calculate total processing time
            finish_time = datetime.now()
            finish_time_str = finish_time.strftime("%Y-%m-%d %H:%M:%S")
            
            if self.start_time:
                total_duration = finish_time - self.start_time
                hours, remainder = divmod(total_duration.total_seconds(), 3600)
                minutes, seconds = divmod(remainder, 60)
                
                if hours > 0:
                    duration_str = f"{int(hours)}h {int(minutes)}m {seconds:.1f}s"
                elif minutes > 0:
                    duration_str = f"{int(minutes)}m {seconds:.1f}s"
                else:
                    duration_str = f"{seconds:.1f}s"
            else:
                duration_str = "Unknown"
            
            self.log_widget.add_log("="*60)
            self.log_widget.add_log(f"🎉 MATCHING PROCESS COMPLETED")
            self.log_widget.add_log(f"⏰ Finish Time: {finish_time_str}")
            self.log_widget.add_log(f"⏱️  Total Duration: {duration_str}")
            self.log_widget.add_log("="*60)
            self.log_widget.add_log("Excel files exported successfully!")
            self.log_widget.add_log("   - Matched files saved to the same folder as input files")
            
            # Mark files as saved and enable output buttons
            self.files_saved = True
            self.results_widget.enable_output_buttons()
            
            QMessageBox.information(self, "Export Complete", f"Matched Excel files have been exported to the same folder as the input files.\n\nTotal processing time: {duration_str}")
            
        except Exception as e:
            # Record finish time even on error
            finish_time = datetime.now()
            finish_time_str = finish_time.strftime("%Y-%m-%d %H:%M:%S")
            
            if self.start_time:
                total_duration = finish_time - self.start_time
                hours, remainder = divmod(total_duration.total_seconds(), 3600)
                minutes, seconds = divmod(remainder, 60)
                
                if hours > 0:
                    duration_str = f"{int(hours)}h {int(minutes)}m {seconds:.1f}s"
                elif minutes > 0:
                    duration_str = f"{int(minutes)}m {seconds:.1f}s"
                else:
                    duration_str = f"{seconds:.1f}s"
            else:
                duration_str = "Unknown"
            
            self.log_widget.add_log("="*60)
            self.log_widget.add_log(f"❌ MATCHING PROCESS FAILED")
            self.log_widget.add_log(f"⏰ Finish Time: {finish_time_str}")
            self.log_widget.add_log(f"⏱️  Total Duration: {duration_str}")
            self.log_widget.add_log("="*60)
            self.log_widget.add_log(f"Export error: {str(e)}")
            QMessageBox.critical(self, "Export Error", f"Failed to export files:\n\n{str(e)}")
        finally:
            # Always reset processing state when done
            self.processing_widget.set_processing_state(False)
    
    def on_matching_error(self, error_message: str):
        """Handle matching errors"""
        # Record finish time even on matching error
        finish_time = datetime.now()
        finish_time_str = finish_time.strftime("%Y-%m-%d %H:%M:%S")
        
        if self.start_time:
            total_duration = finish_time - self.start_time
            hours, remainder = divmod(total_duration.total_seconds(), 3600)
            minutes, seconds = divmod(remainder, 60)
            
            if hours > 0:
                duration_str = f"{int(hours)}h {int(minutes)}m {seconds:.1f}s"
            elif minutes > 0:
                duration_str = f"{int(minutes)}m {seconds:.1f}s"
            else:
                duration_str = f"{seconds:.1f}s"
        else:
            duration_str = "Unknown"
        
        self.log_widget.add_log("="*60)
        self.log_widget.add_log(f"❌ MATCHING PROCESS FAILED")
        self.log_widget.add_log(f"⏰ Finish Time: {finish_time_str}")
        self.log_widget.add_log(f"⏱️  Total Duration: {duration_str}")
        self.log_widget.add_log("="*60)
        self.log_widget.add_log(f"Error: {error_message}")
        
        self.processing_widget.set_processing_state(False)
        QMessageBox.critical(self, "Matching Error", f"An error occurred during matching:\n\n{error_message}")
    
    def open_output_folder(self):
        """Open the input file folder in file explorer"""
        if not self.files_saved:
            QMessageBox.information(self, "Files Not Saved", "Please complete the matching process and save files first.")
            return
            
        if hasattr(self, 'file1_path') and self.file1_path:
            input_dir = Path(self.file1_path).parent
            if input_dir.exists():
                try:
                    os.startfile(str(input_dir))
                except Exception as e:
                    QMessageBox.critical(self, "Error Opening Folder", f"Could not open folder:\n\n{str(e)}")
            else:
                QMessageBox.warning(self, "Folder Not Found", f"Input folder not found: {input_dir}")
        else:
            QMessageBox.information(self, "No Files Selected", "Please select input files first to see the output location.")
    
    def open_output_files(self):
        """Open the output Excel files directly"""
        if not self.files_saved:
            QMessageBox.information(self, "Files Not Saved", "Please complete the matching process and save files first.")
            return
            
        if hasattr(self, 'file1_path') and self.file1_path and hasattr(self, 'file2_path') and self.file2_path:
            try:
                # Get the directory of the input files
                input_dir1 = Path(self.file1_path).parent
                input_dir2 = Path(self.file2_path).parent
                
                # Construct output file paths
                base_name1 = os.path.splitext(os.path.basename(self.file1_path))[0]
                base_name2 = os.path.splitext(os.path.basename(self.file2_path))[0]
                
                output_file1 = input_dir1 / f"{base_name1}_MATCHED.xlsx"
                output_file2 = input_dir2 / f"{base_name2}_MATCHED.xlsx"
                
                # Check if files exist and open them
                files_opened = 0
                if output_file1.exists():
                    os.startfile(str(output_file1))
                    files_opened += 1
                if output_file2.exists():
                    os.startfile(str(output_file2))
                    files_opened += 1
                
                if files_opened == 0:
                    QMessageBox.warning(self, "No Output Files", "No output files found. Please run the matching process first.")
                elif files_opened == 1:
                    QMessageBox.information(self, "Partial Success", "Only one output file was found and opened.")
                # If files_opened == 2, both files were opened successfully (no message needed)
                
            except Exception as e:
                QMessageBox.critical(self, "Error Opening Files", f"Could not open output files:\n\n{str(e)}")
        else:
            QMessageBox.information(self, "No Files Selected", "Please select input files first.")
    
    def closeEvent(self, event):
        """Handle application close event with confirmation"""
        # Always show confirmation dialog for this important program
        if self.matching_thread and self.matching_thread.isRunning():
            reply = QMessageBox.question(
                self,
                "Exit Confirmation",
                "Matching is in progress. Are you sure you want to exit?\n\n"
                "Any unsaved progress will be lost.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.matching_thread.cancel()
                self.matching_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            # Show confirmation even when not processing
            reply = QMessageBox.question(
                self,
                "Exit Confirmation",
                "Are you sure you want to exit the Interunit Loan Matcher?\n\n"
                "Make sure all files have been saved before closing.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                event.accept()
            else:
                event.ignore()
