"""
Bank Account Mapping Widget

Widget for managing bank account name to short code mappings in the GUI.
Allows users to add, edit, and delete mappings with support for up to 4 short codes per account.
"""

import json
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView, QAbstractItemView, QFrame, QSpacerItem, QSizePolicy
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QCursor, QFont

from bank_account_mapping import load_mapping, save_mapping, MAPPING_FILE


class BankAccountMappingWidget(QWidget):
    """Widget for managing bank account mappings"""
    
    mapping_changed = Signal()  # Emitted when mapping is updated
    
    def __init__(self):
        super().__init__()
        self.mapping = load_mapping()
        self.init_ui()
        self.load_mappings_to_table()
    
    def init_ui(self):
        """Initialize the user interface"""
        layout = QVBoxLayout()
        layout.setSpacing(5)
        layout.setContentsMargins(0, 0, 0, 0)  # No outer margins for consistent alignment
        
        # Create section container with consistent padding
        section_container = QWidget()
        section_container.setContentsMargins(0, 0, 0, 0)  # No container margins
        section_layout = QVBoxLayout()
        section_layout.setSpacing(5)  # Consistent spacing between elements
        section_layout.setContentsMargins(15, 15, 15, 10)  # Bottom padding: 10px + 5px spacing = 15px (matches Clear Ledgers button distance)
        section_container.setLayout(section_layout)
        
        # Title
        title = QLabel("Bank Account Mapping")
        title.setProperty("class", "title")
        section_layout.addWidget(title)
        
        # Instructions
        instructions = QLabel(
            "Manage bank account name to short code mappings. Each account can have 1 to 4 short codes. Use comma to separate multiple short codes."
        )
        instructions.setWordWrap(True)
        section_layout.addWidget(instructions)
        
        # Add/Edit section - all in one row
        add_section = QWidget()
        add_layout = QHBoxLayout()
        add_layout.setContentsMargins(0, 0, 0, 0)  # No extra padding, using parent's padding
        add_layout.setSpacing(5)  # Consistent spacing
        
        # Bank Account Name input
        account_label = QLabel("Bank Account Name:")
        add_layout.addWidget(account_label)
        
        self.account_input = QLineEdit()
        self.account_input.setPlaceholderText("e.g., Brac Bank PLC-CD-A/C-2028701210002")
        self.account_input.setMinimumHeight(35)  # Make input box taller
        add_layout.addWidget(self.account_input, 2)  # Stretch factor 2
        
        # Short Codes input
        codes_label = QLabel("Aliases:")
        add_layout.addWidget(codes_label)
        
        self.codes_input = QLineEdit()
        self.codes_input.setPlaceholderText("e.g., BBL#0002, BBL#0003, BBL#0004")
        self.codes_input.setMinimumHeight(35)  # Make input box taller
        add_layout.addWidget(self.codes_input, 2)  # Stretch factor 2
        
        # Add Mapping button
        self.add_button = QPushButton("Add Mapping")
        self.add_button.clicked.connect(self.add_or_update_mapping)
        add_layout.addWidget(self.add_button)
        
        # Clear button
        self.clear_button = QPushButton("Clear")
        self.clear_button.clicked.connect(self.clear_inputs)
        add_layout.addWidget(self.clear_button)
        
        add_section.setLayout(add_layout)
        section_layout.addWidget(add_section)
        
        # Add spacer to match the distance from Clear Ledgers button (10px padding + 5px spacing = 15px total)
        # Since horizontal line is in same layout, we need: 5px (layout spacing) + 10px (spacer) = 15px
        spacer = QSpacerItem(0, 10, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
        section_layout.addItem(spacer)
        
        # Horizontal line separator between input section and table
        horizontal_separator = QFrame()
        horizontal_separator.setFrameShape(QFrame.Shape.HLine)
        horizontal_separator.setFrameShadow(QFrame.Shadow.Sunken)
        horizontal_separator.setStyleSheet("color: #CCCCCC; background-color: #CCCCCC;")
        horizontal_separator.setLineWidth(2)
        horizontal_separator.setFixedHeight(2)
        section_layout.addWidget(horizontal_separator)
        
        # Table section - create container with reduced top padding for equal spacing
        table_section = QWidget()
        table_section.setContentsMargins(0, 0, 0, 0)  # No widget margins
        table_section_layout = QVBoxLayout()
        table_section_layout.setContentsMargins(0, 5, 0, 0)  # Top padding to match Matching tab spacing (5px + 5px layout spacing = 10px)
        table_section_layout.setSpacing(5)
        table_section.setLayout(table_section_layout)
        
        table_label = QLabel("Current Mappings:")
        table_label.setProperty("class", "title")  # Make title bold
        table_section_layout.addWidget(table_label)
        
        # Create table
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Bank Account Name", "Short Codes", "Actions"])
        # Make table headers bold
        header_font = QFont()
        header_font.setBold(True)
        self.table.horizontalHeader().setFont(header_font)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)  # Disable automatic editing
        # Using macOS native table styling (no alternating row colors override)
        # Set row height - no longer need extra space for buttons, using default with padding
        self.table.verticalHeader().setDefaultSectionSize(25)
        table_section_layout.addWidget(self.table)
        
        section_layout.addWidget(table_section)
        
        # Add section container to main layout
        layout.addWidget(section_container)
        
        self.setLayout(layout)
    
    def clear_inputs(self):
        """Clear input fields"""
        self.account_input.clear()
        self.codes_input.clear()
        self.table.clearSelection()
    
    def validate_inputs(self):
        """Validate user inputs"""
        account_name = self.account_input.text().strip()
        codes_text = self.codes_input.text().strip()
        
        if not account_name:
            QMessageBox.warning(self, "Validation Error", "Please enter a bank account name.")
            return False, None, None
        
        if not codes_text:
            QMessageBox.warning(self, "Validation Error", "Please enter at least one short code.")
            return False, None, None
        
        # Parse short codes
        codes = [code.strip() for code in codes_text.split(',') if code.strip()]
        
        if len(codes) == 0:
            QMessageBox.warning(self, "Validation Error", "Please enter at least one short code.")
            return False, None, None
        
        if len(codes) > 4:
            QMessageBox.warning(self, "Validation Error", "Maximum 4 short codes allowed per account.")
            return False, None, None
        
        return True, account_name, codes
    
    def add_or_update_mapping(self):
        """Add or update a bank account mapping"""
        valid, account_name, codes = self.validate_inputs()
        if not valid:
            return
        
        # Add or update the mapping
        self.mapping[account_name] = codes
        
        # Save mappings automatically since the bottom save button was removed
        self.save_mappings()
        
        # Reload table
        self.load_mappings_to_table()
        
        # Clear inputs
        self.clear_inputs()
    
    def delete_mapping(self, account_name):
        """Delete a bank account mapping"""
        reply = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Are you sure you want to delete the mapping for:\n{account_name}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            if account_name in self.mapping:
                del self.mapping[account_name]
                # Save mappings automatically since the bottom save button was removed
                self.save_mappings()
                self.load_mappings_to_table()
    
    def edit_selected_cell(self, row):
        """Make the selected cell in the row editable when Edit button is pressed"""
        # Get the currently selected cell
        current_item = self.table.currentItem()
        
        if current_item and current_item.row() == row:
            # If a cell in this row is selected, make it editable
            self.table.editItem(current_item)
        else:
            # If no cell is selected or selection is in different row, select and edit the account name cell
            account_item = self.table.item(row, 0)
            if account_item:
                self.table.setCurrentItem(account_item)
                self.table.editItem(account_item)
    
    def load_mappings_to_table(self):
        """Load current mappings into the table"""
        self.table.setRowCount(0)
        
        # Disconnect signal temporarily to avoid triggering during load
        try:
            self.table.cellChanged.disconnect()
        except:
            pass
        
        for account_name, codes in sorted(self.mapping.items()):
            row = self.table.rowCount()
            self.table.insertRow(row)
            
            # Account name - editable
            account_item = QTableWidgetItem(account_name)
            account_item.setData(Qt.ItemDataRole.UserRole, account_name)  # Store original name for reference
            self.table.setItem(row, 0, account_item)
            
            # Short codes - editable
            codes_text = ', '.join(codes) if isinstance(codes, list) else str(codes)
            codes_item = QTableWidgetItem(codes_text)
            codes_item.setData(Qt.ItemDataRole.UserRole, account_name)  # Store account name for reference
            self.table.setItem(row, 1, codes_item)
            
            # Actions - using hyperlink-style text
            action_widget = QWidget()
            action_layout = QHBoxLayout()
            action_layout.setContentsMargins(5, 2, 5, 2)
            action_layout.setSpacing(8)  # Optimized spacing for compact view
            
            # Edit link
            edit_link = QLabel("Edit")
            edit_link.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            edit_link.setStyleSheet("color: #1976D2; text-decoration: underline;")
            def edit_clicked(event, row_num=row):
                self.edit_selected_cell(row_num)
            edit_link.mousePressEvent = edit_clicked
            action_layout.addWidget(edit_link)
            
            # Save link
            save_link = QLabel("Save")
            save_link.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            save_link.setStyleSheet("color: #388E3C; text-decoration: underline;")
            def save_clicked(event):
                self.save_mappings()
            save_link.mousePressEvent = save_clicked
            action_layout.addWidget(save_link)
            
            # Delete link
            delete_link = QLabel("Delete")
            delete_link.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            delete_link.setStyleSheet("color: #C62828; text-decoration: underline;")
            def delete_clicked(event, name=account_name):
                self.delete_mapping(name)
            delete_link.mousePressEvent = delete_clicked
            action_layout.addWidget(delete_link)
            
            action_layout.addStretch()
            action_widget.setLayout(action_layout)
            self.table.setCellWidget(row, 2, action_widget)
        
        # Connect cell changed signal to handle direct editing
        self.table.cellChanged.connect(self.on_cell_changed)
    
    def on_cell_changed(self, row, column):
        """Handle direct cell editing in the table"""
        if column == 0:  # Bank Account Name column
            new_account_name = self.table.item(row, column).text().strip()
            old_account_name = self.table.item(row, column).data(Qt.ItemDataRole.UserRole)
            
            if not new_account_name:
                QMessageBox.warning(self, "Validation Error", "Bank account name cannot be empty.")
                # Restore old value
                self.table.item(row, column).setText(old_account_name)
                return
            
            # If account name changed, update the mapping key
            if new_account_name != old_account_name:
                if new_account_name in self.mapping:
                    QMessageBox.warning(self, "Duplicate Account", f"Account '{new_account_name}' already exists.")
                    # Restore old value
                    self.table.item(row, column).setText(old_account_name)
                    return
                
                # Move mapping to new key
                if old_account_name in self.mapping:
                    codes = self.mapping[old_account_name]
                    del self.mapping[old_account_name]
                    self.mapping[new_account_name] = codes
                    # Update user role data
                    self.table.item(row, column).setData(Qt.ItemDataRole.UserRole, new_account_name)
                    # Update codes item reference
                    codes_item = self.table.item(row, 1)
                    if codes_item:
                        codes_item.setData(Qt.ItemDataRole.UserRole, new_account_name)
        
        elif column == 1:  # Short Codes column
            codes_text = self.table.item(row, column).text().strip()
            account_name = self.table.item(row, column).data(Qt.ItemDataRole.UserRole)
            
            if not codes_text:
                QMessageBox.warning(self, "Validation Error", "At least one short code is required.")
                # Restore from mapping
                if account_name in self.mapping:
                    codes = self.mapping[account_name]
                    codes_text = ', '.join(codes) if isinstance(codes, list) else str(codes)
                    self.table.item(row, column).setText(codes_text)
                return
            
            # Parse and validate short codes
            codes = [code.strip() for code in codes_text.split(',') if code.strip()]
            
            if len(codes) == 0:
                QMessageBox.warning(self, "Validation Error", "Please enter at least one short code.")
                # Restore from mapping
                if account_name in self.mapping:
                    codes = self.mapping[account_name]
                    codes_text = ', '.join(codes) if isinstance(codes, list) else str(codes)
                    self.table.item(row, column).setText(codes_text)
                return
            
            if len(codes) > 4:
                QMessageBox.warning(self, "Validation Error", "Maximum 4 short codes allowed per account.")
                # Restore from mapping
                if account_name in self.mapping:
                    codes = self.mapping[account_name]
                    codes_text = ', '.join(codes) if isinstance(codes, list) else str(codes)
                    self.table.item(row, column).setText(codes_text)
                return
            
            # Update mapping
            if account_name in self.mapping:
                self.mapping[account_name] = codes
    
    def save_mappings(self):
        """Save mappings to file"""
        if save_mapping(self.mapping):
            QMessageBox.information(self, "Success", "Bank account mappings saved successfully!")
            self.mapping_changed.emit()
        else:
            QMessageBox.critical(self, "Error", "Failed to save mappings. Please check file permissions.")
    
    def get_mapping(self):
        """Get current mapping dictionary"""
        return self.mapping.copy()
