"""
Matching Logic Package

This package contains all transaction matching algorithms for the Interunit Loan Matcher.

Available Matching Algorithms:
- LCMatchingLogic: Letter of Credit matching
- POMatchingLogic: Purchase Order matching  
- USDMatchingLogic: USD amount matching
- InterunitLoanMatcher: Interunit loan matching
- AggregatedPOMatchingLogic: One-to-many Purchase Order matching (bulk payments)
- NarrationMatchingLogic: Exact text matching (highest priority)
"""

# Import all matching logic classes for convenient access
from .lc_matching_logic import LCMatchingLogic
from .po_matching_logic import POMatchingLogic
from .usd_matching_logic import USDMatchingLogic
from .interunit_loan_matching_logic import InterunitLoanMatcher
from .one_to_many_po_matching_logic import AggregatedPOMatchingLogic
from .narration_matching_logic import NarrationMatchingLogic
from .manual_matching_logic import ManualMatchingLogic
from .final_settlement_matching_logic import FinalSettlementMatchingLogic

# Import patterns for backward compatibility
from .lc_matching_logic import LC_PATTERN
from .po_matching_logic import PO_PATTERN
from .usd_matching_logic import USD_PATTERN

__all__ = [
    'LCMatchingLogic',
    'POMatchingLogic', 
    'USDMatchingLogic',
    'InterunitLoanMatcher',
    'AggregatedPOMatchingLogic',
    'NarrationMatchingLogic',
    'ManualMatchingLogic',
    'FinalSettlementMatchingLogic',
    'LC_PATTERN',
    'PO_PATTERN',
    'USD_PATTERN'
]
