"""
Manual Matching Widget

GUI widget for manually reviewing and confirming/rejecting potential matches
based on amount matching only. Integrated into the main window.
"""

import os
import traceback
from datetime import datetime
from openpyxl import load_workbook
import pandas as pd
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox,
    QGroupBox, QSplitter, QSizePolicy
)
from PySide6.QtCore import Qt, Signal, QTimer
from PySide6.QtGui import QFont, QFontMetrics


class ManualMatchingWidget(QWidget):
    """Widget for manual matching of transaction blocks, integrated into main window."""
    
    # Signals for communication with MainWindow
    finished = Signal(list)  # Emitted with confirmed_matches when all reviewed
    skipped = Signal()       # Emitted when "Skip" is clicked
    cancelled = Signal()     # Emitted if user wants to cancel (if we add a cancel button)

    def __init__(self, parent=None):
        """Initialize manual matching widget."""
        super().__init__(parent)
        self.potential_matches = []
        self.file1_path = ""
        self.file2_path = ""
        self.confirmed_matches = []
        self.current_index = 0
        
        self.ledger_name1 = None
        self.ledger_name2 = None
        self.wb1 = None
        self.wb2 = None
        self.ws1 = None
        self.ws2 = None
        
        self.init_ui()
    
    def load_potential_matches(self, potential_matches, file1_path, file2_path):
        """
        Load new set of potential matches and prepare the widget.
        
        Args:
            potential_matches: List of tuples (block1_rows, block2_rows, amount, file1_is_lender)
            file1_path: Path to File 1 Excel file
            file2_path: Path to File 2 Excel file
        """
        self.potential_matches = potential_matches
        self.file1_path = file1_path
        self.file2_path = file2_path
        self.confirmed_matches = []
        self.current_index = 0
        
        # Extract ledger names
        self.ledger_name1 = self._extract_ledger_name(file1_path)
        self.ledger_name2 = self._extract_ledger_name(file2_path)
        
        # Update group box titles
        file1_title = self.ledger_name1 if self.ledger_name1 else "File 1 Transaction Block"
        self.file1_group.setTitle(file1_title)
        
        file2_title = self.ledger_name2 if self.ledger_name2 else "File 2 Transaction Block"
        self.file2_group.setTitle(file2_title)
        
        # Close old workbooks if they exist
        if self.wb1:
            try: self.wb1.close()
            except: pass
        if self.wb2:
            try: self.wb2.close()
            except: pass
            
        # Load workbooks
        self.wb1 = load_workbook(file1_path, data_only=False)
        self.ws1 = self.wb1.active
        self.wb2 = load_workbook(file2_path, data_only=False)
        self.ws2 = self.wb2.active
        
        # Load first pair
        self.load_match_pair(0)

    def init_ui(self):
        """Initialize the user interface."""
        # Main layout
        main_layout = QVBoxLayout()
        main_layout.setSpacing(10)  # Match container margins for consistent spacing
        main_layout.setContentsMargins(15, 10, 15, 10)  # Consistent with other sections

        # Progress indicator
        progress_label = QLabel()
        progress_label.setProperty("class", "title")  # Use consistent title styling
        progress_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)  # Left align like other titles
        self.progress_label = progress_label
        main_layout.addWidget(progress_label)

        # Match info (amount and types) - reduced spacing
        info_label = QLabel()
        info_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)  # Left align like other titles
        self.info_label = info_label
        main_layout.addWidget(info_label)
        
        # Splitter for side-by-side display - this will expand to fill available space
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # File 1 panel
        self.file1_group = QGroupBox("File 1 Transaction Block")
        group_font = QFont()
        group_font.setBold(True)
        self.file1_group.setFont(group_font)
        self.file1_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        file1_layout = QVBoxLayout()
        file1_layout.setContentsMargins(5, 5, 5, 5)
        file1_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.file1_table = QTableWidget()
        self.file1_table.setAlternatingRowColors(True)
        self.file1_table.horizontalHeader().setStretchLastSection(False)
        self.file1_table.verticalHeader().setVisible(False)
        self.file1_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.file1_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.file1_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        table_font = QFont()
        table_font.setPointSize(8)
        table_font.setBold(False)
        self.file1_table.setFont(table_font)

        header_font = QFont()
        header_font.setPointSize(8)
        header_font.setBold(True)
        self.file1_table.horizontalHeader().setFont(header_font)

        file1_layout.addWidget(self.file1_table, 1)  # Give table stretch factor to expand
        self.file1_group.setLayout(file1_layout)
        splitter.addWidget(self.file1_group)

        # File 2 panel
        self.file2_group = QGroupBox("File 2 Transaction Block")
        self.file2_group.setFont(group_font)
        self.file2_group.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        file2_layout = QVBoxLayout()
        file2_layout.setContentsMargins(5, 5, 5, 5)
        file2_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.file2_table = QTableWidget()
        self.file2_table.setAlternatingRowColors(True)
        self.file2_table.horizontalHeader().setStretchLastSection(False)
        self.file2_table.verticalHeader().setVisible(False)
        self.file2_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.file2_table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.file2_table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self.file2_table.setFont(table_font)
        self.file2_table.horizontalHeader().setFont(header_font)

        file2_layout.addWidget(self.file2_table, 1)  # Give table stretch factor to expand
        self.file2_group.setLayout(file2_layout)
        splitter.addWidget(self.file2_group)

        splitter.setSizes([600, 600])
        splitter.setChildrenCollapsible(False)
        main_layout.addWidget(splitter, 1)  # Give splitter stretch factor 1 to expand

        # Add stretch to push buttons to the bottom
        main_layout.addStretch()

        # Button layout - aligned at bottom
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        # Confirm button - light green
        confirm_button = QPushButton("Confirm Match")
        confirm_button.setProperty("class", "confirm-match-button")
        confirm_button.setMinimumWidth(120)
        confirm_button.setDefault(True)
        confirm_button.clicked.connect(self.on_confirm_match)
        button_layout.addWidget(confirm_button)

        # Reject button - light red
        reject_button = QPushButton("Reject Match")
        reject_button.setProperty("class", "reject-match-button")
        reject_button.setMinimumWidth(120)
        reject_button.clicked.connect(self.on_reject_match)
        button_layout.addWidget(reject_button)

        # Skip button
        skip_button = QPushButton("Skip Manual Matching")
        skip_button.setMinimumWidth(150)
        skip_button.clicked.connect(self.on_skip_manual_matching)
        button_layout.addWidget(skip_button)

        button_layout.addStretch()
        main_layout.addLayout(button_layout)
        
        self.setLayout(main_layout)

        # Ensure the widget expands to fill available space
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
    
    def _extract_ledger_name(self, file_path):
        """Extract ledger name from metadata (row 3, column 0)."""
        try:
            full_df = pd.read_excel(file_path, header=None, nrows=8, dtype=str)
            ledger_name = full_df.iloc[3, 0] if pd.notna(full_df.iloc[3, 0]) else None
            if ledger_name and str(ledger_name).strip() and str(ledger_name).strip().lower() != 'nan':
                return str(ledger_name).strip()
            return None
        except Exception as e:
            print(f"Error extracting ledger name from {os.path.basename(file_path)}: {e}")
            return None
    
    def load_match_pair(self, index):
        """Load and display a match pair."""
        if index < 0 or index >= len(self.potential_matches):
            return
        
        self.current_index = index
        block1_rows, block2_rows, amount, file1_is_lender = self.potential_matches[index]
        
        # Update progress label
        self.progress_label.setText(f"Manual Match Review: {index + 1} of {len(self.potential_matches)}")

        # Update info label
        file1_type = "Lender (Debit)" if file1_is_lender else "Borrower (Credit)"
        file2_type = "Borrower (Credit)" if file1_is_lender else "Lender (Debit)"
        file1_display = self.ledger_name1 if self.ledger_name1 else "File 1"
        file2_display = self.ledger_name2 if self.ledger_name2 else "File 2"
        self.info_label.setText(f"Amount: {amount:,.2f} | {file1_display}: {file1_type} | {file2_display}: {file2_type}")
        
        # Load tables
        self.load_block_to_table(self.file1_table, block1_rows, self.ws1)
        self.load_block_to_table(self.file2_table, block2_rows, self.ws2)
    
    def load_block_to_table(self, table, block_rows, worksheet):
        """Load a transaction block into a table widget. Pixel perfect from original."""
        if not block_rows:
            table.setRowCount(0)
            table.setColumnCount(0)
            return
        
        rows_data = []
        for df_row_idx in block_rows:
            excel_row = df_row_idx + 10
            row_data = []
            for col in range(1, 10):
                cell = worksheet.cell(row=excel_row, column=col)
                value = cell.value
                if value is None:
                    row_data.append("")
                else:
                    if col == 1:
                        if isinstance(value, datetime):
                            row_data.append(value.strftime("%d/%b/%Y"))
                        else:
                            str_value = str(value).strip()
                            if str_value and str_value.lower() not in ['none', 'nan', '']:
                                try:
                                    if '/' in str_value or '-' in str_value:
                                        for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%Y/%m/%d']:
                                            try:
                                                dt = datetime.strptime(str_value, fmt)
                                                row_data.append(dt.strftime("%d/%b/%Y"))
                                                break
                                            except ValueError: continue
                                        else: row_data.append(str_value)
                                    else: row_data.append(str_value)
                                except Exception: row_data.append(str_value)
                            else: row_data.append("")
                    else:
                        row_data.append(str(value))
            rows_data.append(row_data)
        
        num_rows = len(rows_data)
        num_cols = len(rows_data[0]) if rows_data else 9
        
        table.setRowCount(num_rows)
        table.setColumnCount(num_cols)
        
        headers = ["Date", "Dr/Cr", "Particulars", "", "", "Vch Type", "Vch No.", "Debit", "Credit"]
        if num_cols <= len(headers):
            table.setHorizontalHeaderLabels(headers[:num_cols])
        
        for row_idx, row_data in enumerate(rows_data):
            for col_idx, value in enumerate(row_data):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                table.setItem(row_idx, col_idx, item)
        
        if num_cols > 4:
            table.setColumnHidden(3, True)
            table.setColumnHidden(4, True)
        
        table.resizeColumnsToContents()
        if num_cols > 2:
            table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        
        table.updateGeometry()
        
        table_font = table.font()
        font_metrics = QFontMetrics(table_font)
        base_row_height = font_metrics.height() + 4
        
        table.verticalHeader().setDefaultSectionSize(base_row_height)
        table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        for row_idx in range(num_rows):
            table.setRowHeight(row_idx, base_row_height)
        
        table.setWordWrap(True)
        
        def adjust_row_heights():
            particulars_width = table.columnWidth(2) if num_cols > 2 else 0
            if particulars_width > 0:
                for row_idx in range(num_rows):
                    particulars_item = table.item(row_idx, 2)
                    if particulars_item:
                        text = str(particulars_item.text()).strip()
                        if text:
                            text_width = font_metrics.horizontalAdvance(text)
                            if text_width > particulars_width:
                                available_width = particulars_width - 20
                                lines_needed = max(1, (text_width // max(available_width, 1)) + 1)
                                wrapped_height = (font_metrics.height() * lines_needed) + 4
                                table.setRowHeight(row_idx, wrapped_height)
        
        QTimer.singleShot(10, adjust_row_heights)
        
        def set_table_height_to_fit_all_rows():
            header_height = table.horizontalHeader().height()
            if header_height == 0:
                header_height = font_metrics.height() + 8
            total_row_height = 0
            for row_idx in range(num_rows):
                total_row_height += table.rowHeight(row_idx)
            total_height = header_height + total_row_height + 2
            table.setMinimumHeight(total_height)
            table.setMaximumHeight(total_height)
        
        QTimer.singleShot(100, set_table_height_to_fit_all_rows)
    
    def on_confirm_match(self):
        """Handle user confirmation of a match."""
        try:
            if self.current_index >= len(self.potential_matches):
                return
            
            block1_rows, block2_rows, amount, file1_is_lender = self.potential_matches[self.current_index]
            file1_header_row = block1_rows[0]
            file2_header_row = block2_rows[0]
            excel_row1 = file1_header_row + 10
            excel_row2 = file2_header_row + 10
            
            debit_cell1 = self.ws1.cell(row=excel_row1, column=8)
            credit_cell1 = self.ws1.cell(row=excel_row1, column=9)
            file1_debit = float(debit_cell1.value) if debit_cell1.value is not None and debit_cell1.value != 0 else 0.0
            file1_credit = float(credit_cell1.value) if credit_cell1.value is not None and credit_cell1.value != 0 else 0.0
            
            debit_cell2 = self.ws2.cell(row=excel_row2, column=8)
            credit_cell2 = self.ws2.cell(row=excel_row2, column=9)
            file2_debit = float(debit_cell2.value) if debit_cell2.value is not None and debit_cell2.value != 0 else 0.0
            file2_credit = float(credit_cell2.value) if credit_cell2.value is not None and credit_cell2.value != 0 else 0.0
            
            match = {
                'match_id': None,
                'Match_Type': 'Manual',
                'File1_Index': file1_header_row,
                'File2_Index': file2_header_row,
                'File1_Debit': file1_debit,
                'File1_Credit': file1_credit,
                'File2_Debit': file2_debit,
                'File2_Credit': file2_credit,
                'Amount': amount,
                'File1_Amount': file1_debit if file1_debit > 0 else file1_credit,
                'File2_Amount': file2_debit if file2_debit > 0 else file2_credit
            }
            
            self.confirmed_matches.append(match)
            self.move_to_next()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Error confirming match: {str(e)}")
    
    def on_reject_match(self):
        self.move_to_next()
    
    def on_skip_manual_matching(self):
        # Notify MainWindow that user skipped
        if self.wb1: self.wb1.close()
        if self.wb2: self.wb2.close()
        self.skipped.emit()
    
    def move_to_next(self):
        self.current_index += 1
        if self.current_index >= len(self.potential_matches):
            # All matches reviewed
            if self.wb1: self.wb1.close()
            if self.wb2: self.wb2.close()
            self.finished.emit(self.confirmed_matches)
        else:
            self.load_match_pair(self.current_index)
    
    def keyPressEvent(self, event):
        """Prevent accidental Esc closing if the widget has focus."""
        if event.key() == Qt.Key.Key_Escape:
            reply = QMessageBox.question(
                self, "Confirm Cancel",
                "Are you sure you want to stop manual matching?\n\nAny unconfirmed matches will be lost.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                if self.wb1: self.wb1.close()
                if self.wb2: self.wb2.close()
                self.cancelled.emit()
        else:
            super().keyPressEvent(event)
