# Interunit Loan Matcher - GUI Application

A modern PySide6-based graphical user interface for the Interunit Loan Matcher system.

## Features

- **Easy File Selection**: Drag & drop or browse for Excel files
- **Real-time Progress**: Live progress tracking during matching process
- **Results Summary**: Clear overview of matching results by type
- **Export Functionality**: Generate matched Excel files with one click
- **Processing Logs**: Detailed logs of the matching process
- **Modern Interface**: Clean, professional PySide6 interface

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements_gui.txt
```

2. Run the GUI application:
```bash
python run_gui.py
```

Or directly:
```bash
python main_gui.py
```

## Usage

1. **Select Files**: Choose two Excel files using the file selection area
2. **Start Matching**: Click "Start Matching" to begin the automated process
3. **Monitor Progress**: Watch real-time progress for each matching step
4. **View Results**: Review the match summary and statistics
5. **Export Files**: Click "Export Excel Files" to generate matched output files

## Matching Process

The GUI runs an automated matching process with the following steps:

1. **Narration Matching** (Highest Priority)
2. **LC Matching** 
3. **PO Matching**
4. **Interunit Matching**
5. **Salary Matching** (Salary/Remuneration + Month/Year, Festival Bonus + Eid/Year)
6. **Final Settlement Matching**
7. **USD Matching**

See: `project_docs/SALARY_MATCHING_LOGIC.md`

## File Structure

- `main_gui.py` - Main GUI application
- `run_gui.py` - Simple launcher script
- `requirements_gui.txt` - GUI dependencies
- `interunit_loan_matcher.py` - Core matching logic (existing)
- `config.py` - Configuration settings (existing)
- `matching_logic/` - Matching algorithms (existing)

## Requirements

- Python 3.8+
- PySide6
- pandas
- openpyxl
- numpy

## Troubleshooting

If you encounter issues:

1. Ensure all dependencies are installed: `pip install -r requirements_gui.txt`
2. Check that input Excel files are valid and accessible
3. Verify the Output folder exists and is writable
4. Check the processing log for detailed error messages

## Notes

- The GUI provides a user-friendly interface to the existing matching logic
- All matching algorithms run automatically - no user configuration needed
- Output files are saved to the configured Output folder
- The application supports drag & drop file selection for convenience
