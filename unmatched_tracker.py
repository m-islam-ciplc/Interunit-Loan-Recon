"""
Universal Unmatched Record Tracker

This module provides a universal system for tracking why records didn't match
across all matching logic modules. This audit information is written to the
output files' Audit Info column for unmatched records.
"""

from typing import Dict, List, Set
from collections import defaultdict


class UnmatchedTracker:
    """
    Universal tracker for unmatched records and their reasons.
    Used by all matching logic modules to provide audit information.
    """
    
    def __init__(self):
        """Initialize the unmatched tracker."""
        # Dictionary: {file_index: [list of reasons]}
        self.unmatched_reasons: Dict[int, List[str]] = defaultdict(list)
        
        # Track which records were matched (to identify unmatched ones)
        self.matched_indices_file1: Set[int] = set()
        self.matched_indices_file2: Set[int] = set()
    
    def add_unmatched_reason(self, file_index: int, reason: str, file_number: int = 1):
        """
        Add a reason why a record didn't match.
        
        Args:
            file_index: The header row index of the transaction block (can be None)
            reason: Reason why it didn't match (e.g., "Amounts don't match: 1000 vs 2000")
            file_number: 1 for File 1, 2 for File 2
        """
        if file_index is None:
            return  # Skip if index is None
            
        key = f"file{file_number}_index_{file_index}"
        self.unmatched_reasons[key].append(reason)
    
    def mark_as_matched(self, file1_index: int, file2_index: int):
        """
        Mark records as matched so they won't be considered unmatched.
        
        Args:
            file1_index: Header row index from File 1
            file2_index: Header row index from File 2
        """
        self.matched_indices_file1.add(file1_index)
        self.matched_indices_file2.add(file2_index)
    
    def get_unmatched_reasons(self, file_index: int, file_number: int = 1) -> List[str]:
        """
        Get all reasons why a record didn't match.
        
        Args:
            file_index: The header row index of the transaction block
            file_number: 1 for File 1, 2 for File 2
        
        Returns:
            List of reasons, or empty list if no reasons found
        """
        key = f"file{file_number}_index_{file_index}"
        return self.unmatched_reasons.get(key, [])
    
    def get_audit_info_for_unmatched(self, file_index: int, file_number: int = 1) -> str:
        """
        Get formatted audit info string for an unmatched record.
        
        Args:
            file_index: The header row index of the transaction block
            file_number: 1 for File 1, 2 for File 2
        
        Returns:
            Formatted audit info string, or None if record was matched
        """
        # Check if this record was matched
        if file_number == 1 and file_index in self.matched_indices_file1:
            return None
        if file_number == 2 and file_index in self.matched_indices_file2:
            return None
        
        reasons = self.get_unmatched_reasons(file_index, file_number)
        if not reasons:
            return "No match found - No matching criteria met"
        
        # Format reasons into readable audit info
        audit_info = "Unmatched Record\n"
        audit_info += "Reasons:\n"
        for i, reason in enumerate(reasons, 1):
            audit_info += f"{i}. {reason}\n"
        
        return audit_info.strip()
    
    def get_all_unmatched_indices(self, file_number: int = 1) -> Set[int]:
        """
        Get all unmatched indices for a file.
        
        Args:
            file_number: 1 for File 1, 2 for File 2
        
        Returns:
            Set of unmatched header row indices
        """
        if file_number == 1:
            return self.matched_indices_file1
        else:
            return self.matched_indices_file2
    
    def clear(self):
        """Clear all tracked data."""
        self.unmatched_reasons.clear()
        self.matched_indices_file1.clear()
        self.matched_indices_file2.clear()


# Global instance for shared use across all matching modules
_unmatched_tracker_instance = None


def get_unmatched_tracker() -> UnmatchedTracker:
    """
    Get the global unmatched tracker instance.
    This ensures all matching modules use the same tracker.
    
    Returns:
        Global UnmatchedTracker instance
    """
    global _unmatched_tracker_instance
    if _unmatched_tracker_instance is None:
        _unmatched_tracker_instance = UnmatchedTracker()
    return _unmatched_tracker_instance


def reset_unmatched_tracker():
    """Reset the global unmatched tracker (useful for new matching runs)."""
    global _unmatched_tracker_instance
    if _unmatched_tracker_instance is not None:
        _unmatched_tracker_instance.clear()
    _unmatched_tracker_instance = None
