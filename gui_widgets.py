"""
GUI Widgets for Interunit Loan Matcher
Modular UI components for better maintainability
"""

import os
import time
from pathlib import Path
from typing import Dict, Any

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QProgressBar, 
    QTextEdit, QFileDialog, QMessageBox, QApplication, QSpacerItem, QSizePolicy
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent


class FileSelectionWidget(QWidget):
    """Widget for file selection with drag and drop support"""
    
    files_selected = Signal(str, str)  # file1_path, file2_path
    
    def __init__(self):
        super().__init__()
        self.file1_path = ""
        self.file2_path = ""
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(5)
        layout.setContentsMargins(0, 0, 0, 0)  # No outer margins for consistent alignment
        
        # Create section container with curved box
        section_container = QWidget()
        section_container.setContentsMargins(0, 0, 0, 0)  # No container margins
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(15, 15, 15, 10)  # Reduced bottom padding for equal distance from separator line
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Select Interunit Loan Ledgers")
        title.setProperty("class", "title")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        section_layout.addWidget(title)
        
        # Browse button
        self.browse_button = QPushButton("Browse Ledgers")
        self.browse_button.clicked.connect(self.select_both_files)
        section_layout.addWidget(self.browse_button)
        
        # Selected files section
        files_label = QLabel("Selected Ledgers")
        files_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        section_layout.addWidget(files_label)
        
        # File list container - static height for 2 files
        self.files_container = QWidget()
        # Set fixed height to prevent expansion when files are added (2 file labels + padding)
        self.files_container.setFixedHeight(80)  # Fixed height to prevent buttons from moving
        files_layout = QVBoxLayout()
        files_layout.setContentsMargins(0, 0, 0, 0)  # No padding - section already has universal 15px padding
        files_layout.setSpacing(5)  # Consistent spacing
        self.files_container.setLayout(files_layout)
        section_layout.addWidget(self.files_container)
        
        # Add expanding spacer to push buttons down to align with Overall Progress bar
        # This will expand to fill available space, aligning buttons with the progress bar
        spacer = QSpacerItem(0, 0, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding)
        section_layout.addItem(spacer)
        
        # Clear and Run Match buttons side by side
        button_row = QHBoxLayout()
        button_row.setSpacing(5)
        
        # Clear button
        self.clear_files_button = QPushButton("Clear Ledgers")
        self.clear_files_button.clicked.connect(self.clear_files)
        button_row.addWidget(self.clear_files_button)
        
        # Run Match button
        self.run_match_button = QPushButton("Run Match")
        self.run_match_button.clicked.connect(self.run_matching)
        self.run_match_button.setEnabled(False)
        button_row.addWidget(self.run_match_button)
        
        section_layout.addLayout(button_row)
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
        self.setContentsMargins(0, 0, 0, 0)  # No widget-level margins for consistent alignment
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)  # Allow widget to expand
        self.setAcceptDrops(True)
    
    def select_both_files(self):
        """Open file dialog to select both files at once"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Both Excel Files",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if len(files) >= 2:
            # Set first file as File 1 (Pole Book)
            self.set_file(1, files[0])
            # Set second file as File 2 (Steel Book)
            self.set_file(2, files[1])
            
            # If more than 2 files selected, show warning
            if len(files) > 2:
                QMessageBox.information(
                    self, 
                    "Multiple Files Selected", 
                    f"Selected {len(files)} files. Using first two files:\n\n"
                    f"File 1: {os.path.basename(files[0])}\n"
                    f"File 2: {os.path.basename(files[1])}"
                )
        elif len(files) == 1:
            QMessageBox.warning(
                self, 
                "Insufficient Files", 
                "Please select at least 2 Excel files for matching."
            )
    
    def clear_files(self):
        """Clear both file selections"""
        self.file1_path = ""
        self.file2_path = ""
        self.update_file_display()
            
    def set_file(self, file_num: int, file_path: str):
        """Set the selected file path"""
        if file_num == 1:
            self.file1_path = file_path
        else:
            self.file2_path = file_path
        
        self.update_file_display()
        self.validate_files()
    
    def update_file_display(self):
        """Update the file display list"""
        # Clear existing file items
        layout = self.files_container.layout()
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        # Add current files with tick icons and actual filenames
        if self.file1_path:
            file1_name = os.path.basename(self.file1_path)
            file1_item = QLabel(f"✓ {file1_name}")
            file1_item.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            layout.addWidget(file1_item)
        
        if self.file2_path:
            file2_name = os.path.basename(self.file2_path)
            file2_item = QLabel(f"✓ {file2_name}")
            file2_item.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            layout.addWidget(file2_item)
        
        # Enable/disable Run Match button
        self.run_match_button.setEnabled(bool(self.file1_path and self.file2_path))
    
    def set_files(self, file1_path: str, file2_path: str):
        """Set files programmatically (for auto-resume)"""
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.update_file_display()
        self.run_match_button.setEnabled(bool(self.file1_path and self.file2_path))
    
    def run_matching(self):
        """Trigger matching process"""
        if self.file1_path and self.file2_path:
            self.files_selected.emit(self.file1_path, self.file2_path)
            # Emit a signal to start matching
            if hasattr(self.parent(), 'start_matching'):
                self.parent().start_matching()
        
    def validate_files(self):
        """Validate that both files are selected and are valid Excel files"""
        if self.file1_path and self.file2_path:
            # Check if files exist and are Excel files
            if (os.path.exists(self.file1_path) and os.path.exists(self.file2_path) and
                (self.file1_path.endswith('.xlsx') or self.file1_path.endswith('.xls')) and
                (self.file2_path.endswith('.xlsx') or self.file2_path.endswith('.xls'))):
                
                self.files_selected.emit(self.file1_path, self.file2_path)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter event"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        """Handle drop event"""
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        
        # Filter for Excel files only
        excel_files = [f for f in files if f.endswith(('.xlsx', '.xls'))]
        
        if len(excel_files) >= 2:
            self.set_file(1, excel_files[0])
            self.set_file(2, excel_files[1])
            
            # Show info if more than 2 Excel files were dropped
            if len(excel_files) > 2:
                QMessageBox.information(
                    self,
                    "Multiple Files Dropped",
                    f"Dropped {len(excel_files)} Excel files. Using first two:\n\n"
                    f"File 1: {os.path.basename(excel_files[0])}\n"
                    f"File 2: {os.path.basename(excel_files[1])}"
                )
        elif len(excel_files) == 1:
            QMessageBox.warning(
                self,
                "Insufficient Files",
                f"Only 1 Excel file dropped. Please drop at least 2 Excel files.\n\n"
                f"Dropped: {os.path.basename(excel_files[0])}"
            )
        elif len(files) > 0:
            QMessageBox.warning(
                self,
                "Invalid Files",
                f"No Excel files found in dropped files.\n\n"
                f"Please drop .xlsx or .xls files only."
            )
        
        event.acceptProposedAction()


class ProcessingWidget(QWidget):
    """Widget for displaying processing progress and status"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(5)
        layout.setContentsMargins(0, 0, 0, 0)  # No outer margins for consistent alignment
        
        # Create section container with curved box
        section_container = QWidget()
        section_container.setContentsMargins(0, 0, 0, 0)  # No container margins
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(15, 15, 15, 10)  # Reduced bottom padding for equal distance from separator line
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Match Progress")
        title.setProperty("class", "title")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        section_layout.addWidget(title)
        
        # Step progress
        self.step_labels = {}
        self.step_progresses = {}
        
        steps = [
            "Narration Matches",
            "LC Matches",
            "One to One PO Matches",
            "Interunit Matches",
            "Final Settlement Matches",
            "USD Matches"
        ]
        
        for step in steps:
            step_layout = QHBoxLayout()
            step_layout.setSpacing(5)
            
            # Step name
            step_label = QLabel(f"{step}")
            step_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            
            step_layout.addWidget(step_label)
            
            # Step progress bar
            step_progress = QProgressBar()
            step_progress.setRange(0, 100)
            step_progress.setValue(0)
            step_layout.addWidget(step_progress)
            
            section_layout.addLayout(step_layout)
            
            self.step_labels[step] = step_label
            self.step_progresses[step] = step_progress
        
        # Overall progress section - same line layout as step progress bars
        overall_layout = QHBoxLayout()
        overall_layout.setSpacing(5)
        
        overall_label = QLabel("Overall Progress")
        overall_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        overall_layout.addWidget(overall_label)
        
        self.overall_progress = QProgressBar()
        self.overall_progress.setRange(0, 100)
        self.overall_progress.setValue(0)
        overall_layout.addWidget(self.overall_progress)
        
        section_layout.addLayout(overall_layout)
        
        # Add stretch after progress bar to match left widget structure
        section_layout.addStretch()
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
        self.setContentsMargins(0, 0, 0, 0)  # No widget-level margins for consistent alignment
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)  # Allow widget to expand
    
    def complete_step(self, step_name: str, matches_found: int):
        """Mark a step as completed"""
        # Map full step names to display names
        step_mapping = {
            "Narration Matching": "Narration Matches",
            "LC Matching": "LC Matches",
            "PO Matching": "One to One PO Matches",
            "Interunit Matching": "Interunit Matches",
            "Settlement Matching": "Final Settlement Matches",
            "USD Matching": "USD Matches"
        }
        
        short_name = step_mapping.get(step_name, step_name)
        
        if short_name in self.step_progresses:
            progress_bar = self.step_progresses[short_name]
            progress_bar.setValue(100)
            
            # Also update the step label to show completion
            if short_name in self.step_labels:
                pass  # Label updated, no custom styling needed
    
    def reset_progress(self):
        """Reset all progress indicators"""
        for step_name, progress_bar in self.step_progresses.items():
            progress_bar.setValue(0)
            
            # Reset step label styling
            if step_name in self.step_labels:
                pass  # Label reset, no custom styling needed
    
    def set_processing_state(self, is_processing: bool):
        """Enable/disable processing controls"""
        pass


class ResultsWidget(QWidget):
    """Widget for displaying matching results and statistics"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(5)
        
        # Create section container with curved box
        section_container = QWidget()
        section_layout = QVBoxLayout()
        section_layout.setSpacing(5)  # Explicit spacing between elements
        section_layout.setContentsMargins(15, 5, 15, 5)  # Reduced top and bottom padding to move section upward and match spacing (5px + 5px = 10px)
        section_container.setLayout(section_layout)
        
        # Results summary - title and all match types in one row
        single_row_layout = QHBoxLayout()
        single_row_layout.setSpacing(5)
        
        # Title
        title = QLabel("Match Summary")
        title.setProperty("class", "title")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        single_row_layout.addWidget(title)
        
        # Narration Matches
        self.narration_matches_label = QLabel("Narration: 0")
        self.narration_matches_label.setProperty("class", "match-pill narration-pill")
        single_row_layout.addWidget(self.narration_matches_label)
        
        # LC Matches
        self.lc_matches_label = QLabel("LC: 0")
        self.lc_matches_label.setProperty("class", "match-pill lc-pill")
        single_row_layout.addWidget(self.lc_matches_label)
        
        # PO Matches
        self.po_matches_label = QLabel("PO: 0")
        self.po_matches_label.setProperty("class", "match-pill po-pill")
        single_row_layout.addWidget(self.po_matches_label)
        
        # Interunit Matches
        self.interunit_matches_label = QLabel("Interunit: 0")
        self.interunit_matches_label.setProperty("class", "match-pill interunit-pill")
        single_row_layout.addWidget(self.interunit_matches_label)
        
        # USD Matches
        self.usd_matches_label = QLabel("USD: 0")
        self.usd_matches_label.setProperty("class", "match-pill usd-pill")
        single_row_layout.addWidget(self.usd_matches_label)
        
        # Settlement Matches
        self.settlement_matches_label = QLabel("Settlement: 0")
        self.settlement_matches_label.setProperty("class", "match-pill settlement-pill")
        single_row_layout.addWidget(self.settlement_matches_label)
        
        # Total Matches
        single_row_layout.addStretch(1)
        
        self.total_matches_label = QLabel("Total Matches: 0")
        self.total_matches_label.setProperty("class", "match-pill total-pill")
        single_row_layout.addWidget(self.total_matches_label)
        
        section_layout.addLayout(single_row_layout)
        
        # Button row for opening files/folders
        button_row = QHBoxLayout()
        button_row.setSpacing(5)
        
        # Open Folder button
        self.open_folder_button = QPushButton("Open Output Folder")
        # Note: Connection will be set up in MainWindow
        self.open_folder_button.setEnabled(False)  # Disabled until files are processed
        button_row.addWidget(self.open_folder_button)
        
        # Open Output Files button
        self.open_files_button = QPushButton("Open Output Files")
        # Note: Connection will be set up in MainWindow
        self.open_files_button.setEnabled(False)  # Disabled until files are processed
        button_row.addWidget(self.open_files_button)
        
        section_layout.addLayout(button_row)
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
    
    def update_results(self, statistics: Dict[str, Any], enable_buttons: bool = True):
        """Update results display with statistics"""
        self.narration_matches_label.setText(f"Narration: {statistics.get('narration_matches', 0)}")
        self.lc_matches_label.setText(f"LC: {statistics.get('lc_matches', 0)}")
        self.po_matches_label.setText(f"PO: {statistics.get('po_matches', 0)}")
        self.interunit_matches_label.setText(f"Interunit: {statistics.get('interunit_matches', 0)}")
        self.usd_matches_label.setText(f"USD: {statistics.get('usd_matches', 0)}")
        self.settlement_matches_label.setText(f"Settlement: {statistics.get('settlement_matches', 0)}")
        
        # Calculate total matches
        total = (statistics.get('narration_matches', 0) + 
                statistics.get('lc_matches', 0) + 
                statistics.get('po_matches', 0) + 
                statistics.get('interunit_matches', 0) + 
                statistics.get('usd_matches', 0) +
                statistics.get('settlement_matches', 0))
        self.total_matches_label.setText(f"Total Matches: {total}")
        
        # Enable the Open buttons only if explicitly requested and there are matches
        if enable_buttons and total > 0:
            self.open_folder_button.setEnabled(True)
            self.open_files_button.setEnabled(True)
    
    def reset_results(self):
        """Reset results display"""
        self.narration_matches_label.setText("Narration: 0")
        self.lc_matches_label.setText("LC: 0")
        self.po_matches_label.setText("PO: 0")
        self.interunit_matches_label.setText("Interunit: 0")
        self.usd_matches_label.setText("USD: 0")
        self.settlement_matches_label.setText("Settlement: 0")
        self.total_matches_label.setText("Total Matches: 0")
        self.open_folder_button.setEnabled(False)
        self.open_files_button.setEnabled(False)
    
    def enable_output_buttons(self):
        """Enable output buttons after files are saved"""
        self.open_folder_button.setEnabled(True)
        self.open_files_button.setEnabled(True)


class LogWidget(QWidget):
    """Widget for displaying processing logs"""
    
    def __init__(self):
        super().__init__()
        self.init_ui()
        
    def init_ui(self):
        layout = QVBoxLayout()
        layout.setSpacing(0)  # No spacing - section_container is the only child
        layout.setContentsMargins(0, 0, 0, 0)  # No outer margins
        
        # Create section container with curved box
        section_container = QWidget()
        section_container.setContentsMargins(0, 0, 0, 0)  # No container margins
        section_layout = QVBoxLayout()
        section_layout.setContentsMargins(15, 5, 15, 15)  # Reduced top padding to move upward and match spacing (5px + 5px = 10px, same as above Match Summary)
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Process Log")
        title.setProperty("class", "title")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        section_layout.addWidget(title)
        
        # Log text area - expand to fill available space
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        section_layout.addWidget(self.log_text, 1)  # Stretch factor to fill available space
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
        self.setContentsMargins(0, 0, 0, 0)  # No widget-level margins for consistent alignment
    
    def add_log(self, message: str):
        """Add a message to the log with real-time timestamp"""
        # Get timestamp when the message is actually displayed
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        # Auto-scroll to bottom
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
        # Force immediate GUI update
        self.log_text.repaint()
        QApplication.processEvents()
    
    def clear_log(self):
        """Clear the log"""
        self.log_text.clear()
