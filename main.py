"""
Interunit Loan Matcher - GUI Application
Modular PySide6 interface for automated Excel transaction matching
"""

import sys
import os

# Add current directory to Python path for module imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import PySide6 components
from PySide6.QtWidgets import QApplication

def main():
    """Main application entry point"""
    try:
        from main_window import MainWindow
        
        app = QApplication(sys.argv)
        app.setApplicationName("Interunit Loan Matcher")
        app.setApplicationVersion("1.0")
        
        # Set application style to macOS (fallback to Fusion if not available)
        try:
            app.setStyle('macOS')
        except:
            # macOS style not available on this platform, try alternative names
            try:
                app.setStyle('macintosh')
            except:
                # Fallback to Fusion if macOS styles aren't available
                app.setStyle('Fusion')
        
        # Create and show main window
        window = MainWindow()
        window.showMaximized()
        
        # Start event loop
        sys.exit(app.exec())
        
    except ImportError as e:
        print(f"Error importing GUI modules: {e}")
        print("Please install required dependencies:")
        print("pip install -r requirements_gui.txt")
        sys.exit(1)
    except Exception as e:
        print(f"Error starting GUI application: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
