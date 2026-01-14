# Testing Guide for Block Caching Optimization

## Overview
This guide helps you verify that the block caching optimization works correctly and improves performance.

## Test Strategy

### 1. **GUI-Based Testing (Recommended)**

#### Step 1: Baseline Test (Before Optimization)
If you have a backup of the old code, run it first to establish a baseline. Otherwise, skip to Step 2.

1. Use a **large test file** (preferably 10+ MB or 1000+ transactions)
   - Recommended: Files from `Resources/Steel-Pole/` or `Resources/Steel-GeoTex/`
2. Run the matching process
3. **Record the time** shown in the completion message
4. Note the time spent in the "Creating matched Excel files..." phase

#### Step 2: Test with Optimization
1. Use the **same test files** as Step 1
2. Run the matching process
3. **Watch the console/log output** for these messages:
   ```
   === PRE-COMPUTING BLOCK MAPPINGS FOR PERFORMANCE ===
   Pre-computing block mappings for [filename]...
   Block mappings cached successfully!
   === CLEANING UP CACHE ===
   Cache cleared successfully!
   ```
4. **Record the time** and compare with baseline

#### Step 3: Verify Correctness
1. Compare the **output Excel files** from both runs:
   - Same number of matches?
   - Same match IDs?
   - Same transaction blocks identified?
   - Same formatting?

### 2. **Performance Metrics to Check**

#### Expected Improvements:
- **Small files** (< 1 MB, < 100 matches): Minimal improvement (1-2 seconds)
- **Medium files** (1-10 MB, 100-500 matches): Noticeable improvement (5-15 seconds)
- **Large files** (10+ MB, 500+ matches): **Significant improvement** (30+ seconds to minutes)

#### What to Measure:
1. **Total processing time** (shown in completion message)
2. **Time in "Creating matched Excel files..." phase** (this is where caching helps most)
3. **Console output timing** - look for how quickly blocks are retrieved

### 3. **Cache Verification**

#### Check Console Output:
Look for these indicators that caching is working:

**✅ Cache is working if you see:**
```
=== PRE-COMPUTING BLOCK MAPPINGS FOR PERFORMANCE ===
Pre-computing block mappings for File1.xlsx...
Pre-computing block mappings for File2.xlsx...
Block mappings cached successfully!
```

**❌ Cache is NOT working if:**
- You see repeated "Loading workbook..." messages during file creation
- The "Creating matched Excel files..." phase takes a very long time
- No "PRE-COMPUTING BLOCK MAPPINGS" message appears

### 4. **Edge Case Testing**

#### Test Case 1: Small Files
- **Purpose**: Verify caching doesn't break small file processing
- **Files**: Use smallest available test files
- **Expected**: Should work correctly, minimal performance gain

#### Test Case 2: Very Large Files
- **Purpose**: Verify caching handles large datasets
- **Files**: Use largest available test files
- **Expected**: Significant performance improvement

#### Test Case 3: Files with Many Matches
- **Purpose**: Verify caching works when many matches are found
- **Files**: Files known to have many matches
- **Expected**: Faster file creation phase

#### Test Case 4: Files with Few Matches
- **Purpose**: Verify caching works even with few matches
- **Files**: Files with minimal matches
- **Expected**: Should work correctly

### 5. **Correctness Verification Checklist**

After running tests, verify:

- [ ] **Match Count**: Same number of matches found
- [ ] **Match IDs**: Sequential and unique (1, 2, 3, ...)
- [ ] **Match Types**: Same match types assigned
- [ ] **Transaction Blocks**: All blocks correctly identified
- [ ] **Output Format**: Excel files formatted correctly
- [ ] **Unmatched Records**: All unmatched records included
- [ ] **Audit Info**: Audit information correctly placed

### 6. **Quick Test Script**

Run the provided `test_block_caching.py` script for automated testing:

```bash
python test_block_caching.py
```

This script will:
- Test cache functionality directly
- Measure performance improvements
- Verify correctness

## Troubleshooting

### Issue: No performance improvement
**Possible causes:**
1. Files are too small (< 100 matches)
2. Cache not being used (check console output)
3. Other bottlenecks (file I/O, Excel formatting)

**Solution:** Use larger test files (10+ MB)

### Issue: Different results than before
**Possible causes:**
1. Cache returning incorrect block mappings
2. Logic error in cache implementation

**Solution:** Compare output files carefully, check console logs

### Issue: Memory errors
**Possible causes:**
1. Very large files causing memory issues
2. Cache not being cleared properly

**Solution:** Check that `clear_cache()` is being called

## Success Criteria

✅ **Performance**: 30%+ improvement for large files (10+ MB, 500+ matches)
✅ **Correctness**: Identical results to previous version
✅ **Stability**: No crashes or errors
✅ **Cache**: Console shows cache being used

## Reporting Results

When reporting test results, include:
1. File sizes (MB)
2. Number of transactions
3. Number of matches found
4. Processing time (before/after if available)
5. Any errors or issues encountered
