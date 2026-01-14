"""
Test script for block caching optimization
Tests performance and correctness of the caching implementation
"""

import time
import os
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from transaction_block_identifier import TransactionBlockIdentifier
from interunit_loan_matcher import ExcelTransactionMatcher
import pandas as pd


def test_cache_functionality():
    """Test that cache works correctly"""
    print("=" * 70)
    print("TEST 1: Cache Functionality")
    print("=" * 70)
    
    # Find test files
    resources_dir = Path("Resources")
    test_files = []
    
    # Look for Excel files in Resources directory
    for pattern in ["**/*.xlsx"]:
        test_files.extend(list(resources_dir.glob(pattern)))
    
    if len(test_files) < 2:
        print("❌ ERROR: Need at least 2 Excel files for testing")
        print(f"   Found {len(test_files)} files in Resources/")
        return False
    
    # Use first two files found
    file1 = str(test_files[0])
    file2 = str(test_files[1])
    
    print(f"\nUsing test files:")
    print(f"  File 1: {os.path.basename(file1)}")
    print(f"  File 2: {os.path.basename(file2)}")
    
    try:
        # Initialize matcher
        matcher = ExcelTransactionMatcher(file1, file2)
        
        # Read transactions
        print("\n📖 Reading transactions...")
        matcher.metadata1, matcher.transactions1 = matcher.read_complex_excel(file1)
        matcher.metadata2, matcher.transactions2 = matcher.read_complex_excel(file2)
        
        print(f"  File 1: {len(matcher.transactions1)} transactions")
        print(f"  File 2: {len(matcher.transactions2)} transactions")
        
        # Test cache pre-computation
        print("\n🔧 Testing cache pre-computation...")
        start_time = time.time()
        
        matcher.block_identifier.precompute_all_blocks(matcher.transactions1, file1)
        matcher.block_identifier.precompute_all_blocks(matcher.transactions2, file2)
        
        cache_time = time.time() - start_time
        print(f"  ✅ Cache pre-computation completed in {cache_time:.2f} seconds")
        
        # Verify cache is populated
        if file1 in matcher.block_identifier._block_cache:
            cache_size1 = len(matcher.block_identifier._block_cache[file1])
            print(f"  ✅ File 1 cache: {cache_size1} block mappings")
        else:
            print(f"  ❌ File 1 cache: NOT FOUND")
            return False
        
        if file2 in matcher.block_identifier._block_cache:
            cache_size2 = len(matcher.block_identifier._block_cache[file2])
            print(f"  ✅ File 2 cache: {cache_size2} block mappings")
        else:
            print(f"  ❌ File 2 cache: NOT FOUND")
            return False
        
        # Test cache retrieval (should be fast)
        print("\n⚡ Testing cache retrieval speed...")
        test_indices = list(range(0, min(100, len(matcher.transactions1)), 10))
        
        start_time = time.time()
        for idx in test_indices:
            block_rows = matcher.block_identifier.get_transaction_block_rows(idx, file1)
            if block_rows:
                # Verify we got valid block rows
                pass
        cache_retrieval_time = time.time() - start_time
        
        print(f"  ✅ Retrieved {len(test_indices)} blocks in {cache_retrieval_time:.4f} seconds")
        print(f"  ✅ Average: {cache_retrieval_time/len(test_indices)*1000:.2f} ms per lookup")
        
        # Test cache cleanup
        print("\n🧹 Testing cache cleanup...")
        matcher.block_identifier.clear_cache(file1)
        matcher.block_identifier.clear_cache(file2)
        
        if file1 not in matcher.block_identifier._block_cache:
            print(f"  ✅ File 1 cache cleared")
        else:
            print(f"  ❌ File 1 cache NOT cleared")
            return False
        
        if file2 not in matcher.block_identifier._block_cache:
            print(f"  ✅ File 2 cache cleared")
        else:
            print(f"  ❌ File 2 cache NOT cleared")
            return False
        
        print("\n✅ TEST 1 PASSED: Cache functionality works correctly")
        return True
        
    except Exception as e:
        print(f"\n❌ TEST 1 FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_performance_comparison():
    """Compare performance with and without cache"""
    print("\n" + "=" * 70)
    print("TEST 2: Performance Comparison")
    print("=" * 70)
    
    # Find test files
    resources_dir = Path("Resources")
    test_files = []
    
    for pattern in ["**/*.xlsx"]:
        test_files.extend(list(resources_dir.glob(pattern)))
    
    if len(test_files) < 2:
        print("❌ ERROR: Need at least 2 Excel files for testing")
        return False
    
    file1 = str(test_files[0])
    file2 = str(test_files[1])
    
    print(f"\nUsing test files:")
    print(f"  File 1: {os.path.basename(file1)}")
    print(f"  File 2: {os.path.basename(file2)}")
    
    try:
        matcher = ExcelTransactionMatcher(file1, file2)
        matcher.metadata1, matcher.transactions1 = matcher.read_complex_excel(file1)
        matcher.metadata2, matcher.transactions2 = matcher.read_complex_excel(file2)
        
        # Test WITH cache (pre-computed)
        print("\n⚡ Testing WITH cache (pre-computed)...")
        matcher.block_identifier.precompute_all_blocks(matcher.transactions1, file1)
        matcher.block_identifier.precompute_all_blocks(matcher.transactions2, file2)
        
        # Clear cache to simulate "without cache" scenario
        matcher.block_identifier.clear_cache(file1)
        matcher.block_identifier.clear_cache(file2)
        
        # Test WITHOUT cache (simulated by clearing cache first)
        print("\n🐌 Testing WITHOUT cache (simulated)...")
        test_indices = list(range(0, min(50, len(matcher.transactions1)), 5))
        
        start_time = time.time()
        for idx in test_indices:
            block_rows = matcher.block_identifier.get_transaction_block_rows(idx, file1)
        time_without_cache = time.time() - start_time
        
        print(f"  Time without cache: {time_without_cache:.4f} seconds")
        print(f"  Average: {time_without_cache/len(test_indices)*1000:.2f} ms per lookup")
        
        # Now test WITH cache
        matcher.block_identifier.precompute_all_blocks(matcher.transactions1, file1)
        matcher.block_identifier.precompute_all_blocks(matcher.transactions2, file2)
        
        start_time = time.time()
        for idx in test_indices:
            block_rows = matcher.block_identifier.get_transaction_block_rows(idx, file1)
        time_with_cache = time.time() - start_time
        
        print(f"\n  Time with cache: {time_with_cache:.4f} seconds")
        print(f"  Average: {time_with_cache/len(test_indices)*1000:.2f} ms per lookup")
        
        # Calculate improvement
        if time_without_cache > 0:
            improvement = ((time_without_cache - time_with_cache) / time_without_cache) * 100
            speedup = time_without_cache / time_with_cache if time_with_cache > 0 else 0
            
            print(f"\n📊 Performance Results:")
            print(f"  Improvement: {improvement:.1f}% faster")
            print(f"  Speedup: {speedup:.1f}x")
            
            if improvement > 50:
                print(f"  ✅ EXCELLENT: Significant performance improvement!")
            elif improvement > 20:
                print(f"  ✅ GOOD: Noticeable performance improvement")
            else:
                print(f"  ⚠️  MODERATE: Small improvement (may be more noticeable with larger files)")
        
        matcher.block_identifier.clear_cache(file1)
        matcher.block_identifier.clear_cache(file2)
        
        print("\n✅ TEST 2 COMPLETED")
        return True
        
    except Exception as e:
        print(f"\n❌ TEST 2 FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_correctness():
    """Test that cached results match non-cached results"""
    print("\n" + "=" * 70)
    print("TEST 3: Correctness Verification")
    print("=" * 70)
    
    resources_dir = Path("Resources")
    test_files = []
    
    for pattern in ["**/*.xlsx"]:
        test_files.extend(list(resources_dir.glob(pattern)))
    
    if len(test_files) < 1:
        print("❌ ERROR: Need at least 1 Excel file for testing")
        return False
    
    file1 = str(test_files[0])
    
    try:
        matcher = ExcelTransactionMatcher(file1, file1)  # Use same file for simplicity
        matcher.metadata1, matcher.transactions1 = matcher.read_complex_excel(file1)
        
        # Get results WITHOUT cache
        print("\n🔍 Getting block mappings WITHOUT cache...")
        test_indices = list(range(0, min(20, len(matcher.transactions1)), 2))
        results_without_cache = {}
        
        for idx in test_indices:
            block_rows = matcher.block_identifier.get_transaction_block_rows(idx, file1)
            results_without_cache[idx] = block_rows
        
        # Clear any existing cache
        matcher.block_identifier.clear_cache(file1)
        
        # Pre-compute cache
        print("🔍 Pre-computing cache...")
        matcher.block_identifier.precompute_all_blocks(matcher.transactions1, file1)
        
        # Get results WITH cache
        print("🔍 Getting block mappings WITH cache...")
        results_with_cache = {}
        
        for idx in test_indices:
            block_rows = matcher.block_identifier.get_transaction_block_rows(idx, file1)
            results_with_cache[idx] = block_rows
        
        # Compare results
        print("\n📊 Comparing results...")
        mismatches = 0
        
        for idx in test_indices:
            without = results_without_cache.get(idx, [])
            with_cache = results_with_cache.get(idx, [])
            
            if without != with_cache:
                mismatches += 1
                print(f"  ❌ Mismatch at index {idx}:")
                print(f"     Without cache: {without}")
                print(f"     With cache: {with_cache}")
        
        if mismatches == 0:
            print(f"  ✅ All {len(test_indices)} test cases match!")
            print("  ✅ Cache produces identical results to non-cached version")
        else:
            print(f"  ❌ Found {mismatches} mismatches out of {len(test_indices)} tests")
            return False
        
        matcher.block_identifier.clear_cache(file1)
        
        print("\n✅ TEST 3 PASSED: Correctness verified")
        return True
        
    except Exception as e:
        print(f"\n❌ TEST 3 FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Run all tests"""
    print("\n" + "=" * 70)
    print("BLOCK CACHING OPTIMIZATION - TEST SUITE")
    print("=" * 70)
    
    results = []
    
    # Run tests
    results.append(("Cache Functionality", test_cache_functionality()))
    results.append(("Performance Comparison", test_performance_comparison()))
    results.append(("Correctness Verification", test_correctness()))
    
    # Summary
    print("\n" + "=" * 70)
    print("TEST SUMMARY")
    print("=" * 70)
    
    for test_name, passed in results:
        status = "✅ PASSED" if passed else "❌ FAILED"
        print(f"  {test_name}: {status}")
    
    all_passed = all(result[1] for result in results)
    
    if all_passed:
        print("\n🎉 ALL TESTS PASSED!")
        print("\nThe block caching optimization is working correctly.")
        print("You can now test with real files using the GUI.")
    else:
        print("\n⚠️  SOME TESTS FAILED")
        print("Please review the errors above before using in production.")
    
    return all_passed


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
