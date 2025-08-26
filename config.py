# =============================================================================
# CONFIGURATION MODULE
# =============================================================================
# Centralized configuration for all regex patterns and system settings

# Regex patterns have been moved to their respective modules:
# - LC_PATTERN → lc_matching_logic.py
# - PO_PATTERN → po_matching_logic.py

# Amount matching tolerance (for rounding differences)
# AMOUNT_TOLERANCE = 0.01  # ❌ UNUSED - removed since all matching uses exact amounts

# File paths and output settings
# INPUT_FILE1_PATH = "Input Files/Interunit Steel.xlsx"
# INPUT_FILE2_PATH = "Input Files/Interunit GeoTex.xlsx"

INPUT_FILE1_PATH = "Input Files/Pole Book STEEL.xlsx"
INPUT_FILE2_PATH = "Input Files/Steel Book POLE.xlsx"

# INPUT_FILE1_PATH = "Input Files/Steel book Trans.xlsx"
# INPUT_FILE2_PATH = "Input Files/Trans book Steel.xlsx"

OUTPUT_FOLDER = "Output"
OUTPUT_SUFFIX = "_MATCHED.xlsx"
SIMPLE_SUFFIX = "_SIMPLE.xlsx"
CREATE_SIMPLE_FILES = False
# CREATE_ALT_FILES = False  # ❌ UNUSED - commenting out
VERBOSE_DEBUG = True

def print_configuration():
    """Print current configuration settings."""
    print("=" * 60)
    print("CURRENT CONFIGURATION")
    print("=" * 60)
    print(f"Input File 1: {INPUT_FILE1_PATH}")
    print(f"Input File 2: {INPUT_FILE2_PATH}")
    print(f"Output Folder: {OUTPUT_FOLDER}")
    print(f"Output Suffix: {OUTPUT_SUFFIX}")
    print(f"Simple Files: {'Yes' if CREATE_SIMPLE_FILES else 'No'}")
    # print(f"Alternative Files: {'Yes' if CREATE_ALT_FILES else 'No'}")  # ❌ UNUSED - commenting out
    print(f"Verbose Debug: {'Yes' if VERBOSE_DEBUG else 'No'}")
    print("LC Pattern: Defined in lc_matching_logic.py")
    print("PO Pattern: Defined in po_matching_logic.py")
    # print(f"Amount Tolerance: {AMOUNT_TOLERANCE}")  # ❌ UNUSED - removed
    print("=" * 60)

def update_configuration():
    """Interactive configuration update (for future use)."""
    print("To update configuration, modify the variables in config.py:")
    print("1. INPUT_FILE1_PATH - Path to your first Excel file")
    print("2. INPUT_FILE2_PATH - Path to your second Excel file")
    print("3. OUTPUT_FOLDER - Where to save output files")
    print("4. OUTPUT_SUFFIX - Suffix for matched files")
    print("5. CREATE_SIMPLE_FILES - Whether to create simple test files")
    # print("6. CREATE_ALT_FILES - Whether to create alternative files")  # ❌ UNUSED - commenting out
    print("7. VERBOSE_DEBUG - Whether to show detailed debug output")
    print("8. LC_PATTERN - Defined in lc_matching_logic.py")
    print("9. PO_PATTERN - Defined in po_matching_logic.py")
    # print("10. AMOUNT_TOLERANCE - Tolerance for amount matching (0 for exact)")  # ❌ UNUSED - removed
