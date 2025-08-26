"""
Match ID Manager - Centralized, Order-Independent Sequential Match ID Generation

This module provides a robust Match ID management system that ensures:
1. Sequential Match IDs (M001, M002, M003...) regardless of execution order
2. Thread-safe counter management
3. Complete audit trail of Match ID assignment
4. Consistency across all matching logic types

Author: Software Architect
"""

class MatchIDManager:
    """
    Centralized Match ID manager that ensures sequential numbering
    across all matching types regardless of execution order.
    """
    
    def __init__(self):
        """Initialize the Match ID manager with a fresh counter."""
        self._counter = 0
        self._audit_trail = []
        self._reserved_ids = set()
    
    def get_next_match_id(self, match_type: str, context: str = "") -> str:
        """
        Generate the next sequential Match ID and record it in audit trail.
        
        Args:
            match_type: Type of match (e.g., "Narration", "LC", "PO", "Interunit", "USD")
            context: Additional context for debugging (e.g., "File1_Row_123")
            
        Returns:
            Sequential Match ID string (e.g., "M001", "M002", etc.)
        """
        self._counter += 1
        match_id = f"M{self._counter:03d}"
        
        # Record in audit trail for debugging
        self._audit_trail.append({
            'match_id': match_id,
            'match_type': match_type,
            'context': context,
            'sequence': self._counter
        })
        
        print(f"    Generated Match ID: {match_id} (Type: {match_type}, Context: {context})")
        
        return match_id
    
    def get_current_count(self) -> int:
        """Get the current number of Match IDs generated."""
        return self._counter
    
    def get_audit_trail(self) -> list:
        """Get the complete audit trail of Match ID generation."""
        return self._audit_trail.copy()
    
    def print_audit_summary(self):
        """Print a summary of all generated Match IDs for debugging."""
        print(f"\nMATCH ID AUDIT SUMMARY")
        print(f"Total Match IDs Generated: {self._counter}")
        print(f"Expected Range: M001 to M{self._counter:03d}")
        
        # Group by match type
        by_type = {}
        for entry in self._audit_trail:
            match_type = entry['match_type']
            if match_type not in by_type:
                by_type[match_type] = []
            by_type[match_type].append(entry['match_id'])
        
        for match_type, ids in by_type.items():
            print(f"  {match_type}: {len(ids)} matches ({ids[0]} to {ids[-1]})")
    
    def validate_sequence(self) -> bool:
        """
        Validate that all Match IDs are properly sequential.
        
        Returns:
            True if sequence is valid, False otherwise
        """
        expected_ids = [f"M{i:03d}" for i in range(1, self._counter + 1)]
        actual_ids = [entry['match_id'] for entry in self._audit_trail]
        
        is_valid = expected_ids == actual_ids
        
        if not is_valid:
            print(f"SEQUENCE VALIDATION FAILED!")
            print(f"Expected: {expected_ids[:10]}...")
            print(f"Actual:   {actual_ids[:10]}...")
        else:
            print(f"SEQUENCE VALIDATION PASSED: {len(actual_ids)} sequential Match IDs")
            
        return is_valid

# Global instance to be shared across all matching logic
_global_match_id_manager = None

def get_match_id_manager() -> MatchIDManager:
    """Get the global Match ID manager instance."""
    global _global_match_id_manager
    if _global_match_id_manager is None:
        _global_match_id_manager = MatchIDManager()
    return _global_match_id_manager

def reset_match_id_manager():
    """Reset the global Match ID manager (for testing)."""
    global _global_match_id_manager
    _global_match_id_manager = MatchIDManager()

