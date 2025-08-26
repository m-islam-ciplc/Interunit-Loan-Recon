"""
Bank Account Name to Short Code Mapping

This module contains the mapping between full bank account names and their
corresponding short codes used in interunit loan matching.

Format: Full Account Name → List of Short Codes (up to 4)
Example: 'Brac Bank PLC-CD-A/C-2028701210002' → ['BBL#0002']
"""

import json
import os

# Default interunit account mapping (Full Format → List of Short Codes)
# Each account can have 1-4 short codes
DEFAULT_INTERUNIT_ACCOUNT_MAPPING = {
    'BRAC Bank Ltd.-STD-1540102870121001': ['BBL#1001', 'BBL#21001'],
    'Brac Bank PLC-CD-A/C-2028701210002': ['BBL#0002', 'BBL#10002'],
    'Dhaka Bank OD-A/C-2051750000205-CIL': ['DBL#0205', 'DBL#00205'],
    'Dhaka Bank-STD-2051501833-CIL': ['DBL#1833', 'DBL#01833'],
    'Dutch Bangla Bank Ltd.-SND-1071200003988': ['DBBL#3988', 'DBBL#03988'],
    'Eastern Bank Limited-SND-1011060605503': ['EBL#5503', 'EBL#05503'],
    'Eastern Bank OD#1012040163265': ['EBL#3265', 'EBL#63265'],
    'Eastern Bank OD#1012210603129': ['EBL#3129', 'EBL#03129'],
    'Eastern Bank,STD-1011220144056': ['EBL#4056', 'EBL#44056'],
    'MTBL-SND-A/C-1310000003858': ['MTBL#3858', 'MTBL#03858'],
    'Midland Bank Ltd-CE-0011-1060000313': ['MDB#0313', 'MDB#00313'],
    'Midland Bank PLC-CD-A/C-0011-1050011026': ['MDB#11026', 'MDB#1026'],
    'Midland-CE-0011-1060000304-CI': ['MDB#0304', 'MDB#00304'],
    'Midland-CE-0011-1060000331-CI': ['MDB#0331', 'MDB#00331'],
    'One Bank-CD/A/C-0011020008826': ['OBL#8826', 'OBL#08826'],
    'One Bank-SND-A/C-0013000002451': ['OBL#2451', 'OBL#02451'],
    'PBL-SND- 2126312011060': ['PBL#11060', 'PBL#1060'],
    'Prime Bank Limited-SND-2126318011502': ['PBL#11502', 'PBL#1502'],
    'Prime Bank-CD-2126117010855': ['PBL#10855', 'PBL#0855'],
}


# File path for saving/loading mappings
MAPPING_FILE = 'bank_account_mapping.json'

def load_mapping():
    """Load bank account mapping from Python file (single source of truth)."""
    # Python file is the single source of truth
    return DEFAULT_INTERUNIT_ACCOUNT_MAPPING.copy()

def save_mapping(mapping):
    """Save bank account mapping to Python file (single source of truth)."""
    try:
        # Read the current Python file
        py_file_path = __file__
        with open(py_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Find the DEFAULT_INTERUNIT_ACCOUNT_MAPPING dictionary
        start_marker = 'DEFAULT_INTERUNIT_ACCOUNT_MAPPING = {'
        end_marker = '}'
        
        start_idx = content.find(start_marker)
        if start_idx == -1:
            raise ValueError("Could not find DEFAULT_INTERUNIT_ACCOUNT_MAPPING in file")
        
        # Find the matching closing brace
        brace_count = 0
        end_idx = start_idx
        for i in range(start_idx, len(content)):
            if content[i] == '{':
                brace_count += 1
            elif content[i] == '}':
                brace_count -= 1
                if brace_count == 0:
                    end_idx = i
                    break
        
        if brace_count != 0:
            raise ValueError("Could not find matching closing brace for DEFAULT_INTERUNIT_ACCOUNT_MAPPING")
        
        # Get the line before the dict to preserve formatting
        line_before = content[:start_idx].split('\n')[-1]
        indent_level = len(line_before) - len(line_before.lstrip())
        entry_indent = ' ' * (indent_level + 4)  # 4 spaces for dict entries
        
        # Generate new dictionary content
        new_dict_content = 'DEFAULT_INTERUNIT_ACCOUNT_MAPPING = {\n'
        
        # Sort keys for consistency
        for key, value in sorted(mapping.items()):
            # Format the entry - escape single quotes in keys if needed
            escaped_key = key.replace("'", "\\'")
            if isinstance(value, list):
                value_str = str(value)
            else:
                value_str = f"['{value}']"
            new_dict_content += f"{entry_indent}'{escaped_key}': {value_str},\n"
        
        new_dict_content += ' ' * indent_level + '}\n'
        
        # Replace the dictionary section
        new_content = content[:start_idx] + new_dict_content + content[end_idx + 1:]
        
        # Write back to file
        with open(py_file_path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        
        # Update the module-level variable
        global DEFAULT_INTERUNIT_ACCOUNT_MAPPING, INTERUNIT_ACCOUNT_MAPPING
        DEFAULT_INTERUNIT_ACCOUNT_MAPPING = mapping.copy()
        INTERUNIT_ACCOUNT_MAPPING = mapping.copy()
        
        return True
    except Exception as e:
        print(f"Error saving mapping to Python file: {e}")
        import traceback
        traceback.print_exc()
        return False

# Load the mapping on import
INTERUNIT_ACCOUNT_MAPPING = load_mapping()
