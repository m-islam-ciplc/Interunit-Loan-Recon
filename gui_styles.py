"""
GUI Styles for Interunit Loan Matcher
Centralized CSS styling for better maintainability
"""

def get_main_stylesheet():
    """Get the main application stylesheet - minimal styling for button bold text and titles"""
    return """
        QPushButton {
            font-weight: bold;
            min-height: 25px;
        }
        QPushButton[class="run-match-button"] {
            background-color: #E3F2FD;
            color: #1976D2;
            border: 1px solid #BBDEFB;
            border-radius: 4px;
        }
        QPushButton[class="run-match-button"]:hover {
            background-color: #BBDEFB;
        }
        QPushButton[class="run-match-button"]:pressed {
            background-color: #90CAF9;
        }
        /* Manual Matching Buttons */
        QPushButton[class="confirm-match-button"] {
            background-color: #E8F5E9;
            color: #388E3C;
            border: 1px solid #C8E6C9;
            border-radius: 4px;
        }
        QPushButton[class="confirm-match-button"]:hover {
            background-color: #C8E6C9;
        }
        QPushButton[class="confirm-match-button"]:pressed {
            background-color: #A5D6A7;
        }
        QPushButton[class="reject-match-button"] {
            background-color: #FFEBEE;
            color: #C62828;
            border: 1px solid #FFCDD2;
            border-radius: 4px;
        }
        QPushButton[class="reject-match-button"]:hover {
            background-color: #FFCDD2;
        }
        QPushButton[class="reject-match-button"]:pressed {
            background-color: #EF9A9A;
        }
        QLabel[class="title"] {
            font-weight: bold;
        }
        /* Match Summary Pills/Tags - Button-like style */
        QLabel[class*="match-pill"] {
            border-radius: 4px;
            padding: 4px 10px;
            font-size: 12px;
            font-weight: 500;
            border: 1px solid;
        }
        QLabel[class*="narration-pill"] {
            background-color: #E3F2FD;
            color: #1976D2;
            border-color: #BBDEFB;
        }
        QLabel[class*="lc-pill"] {
            background-color: #E8F5E9;
            color: #388E3C;
            border-color: #C8E6C9;
        }
        QLabel[class*="po-pill"] {
            background-color: #FFF3E0;
            color: #F57C00;
            border-color: #FFE0B2;
        }
        QLabel[class*="interunit-pill"] {
            background-color: #F3E5F5;
            color: #7B1FA2;
            border-color: #E1BEE7;
        }
        QLabel[class*="usd-pill"] {
            background-color: #E0F7FA;
            color: #00838F;
            border-color: #B2EBF2;
        }
        QLabel[class*="settlement-pill"] {
            background-color: #FFEBEE;
            color: #C62828;
            border-color: #FFCDD2;
        }
        QLabel[class*="total-pill"] {
            background-color: #F5F5F5;
            color: #424242;
            border-color: #E0E0E0;
            font-weight: bold;
        }
        /* Table Selection Styling */
        QTableWidget {
            selection-background-color: #BBDEFB;
            selection-color: #000000;
            outline: none;
        }
        QTableWidget::item:hover {
            background-color: #E3F2FD;
        }
        QTableWidget::item:selected {
            background-color: #90CAF9;
            color: #000000;
        }
    """
