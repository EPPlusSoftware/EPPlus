/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/09/2026         EPPlus Software AB       RangeCriteria performance optimization
 *************************************************************************************************/
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions
{
    /// <summary>
    /// Cache for RangeCriteria functions (SumIfs, AverageIfs, CountIfs) to improve performance
    /// during calculation by avoiding repeated expensive operations.
    /// Uses FIFO eviction policy to limit memory usage and integer keys for efficient lookups.
    /// </summary>
    internal class RangeCriteriaCache
    {
        private const int DEFAULT_MAX_FLATTENED_RANGES = 100;
        private const int DEFAULT_MAX_MATCH_INDEXES = 1000;

        private readonly int _maxFlattenedRanges;
        private readonly int _maxMatchIndexes;
        private readonly ExcelPackage _package;

        private Dictionary<int, List<object>> _flattenedRanges;
        private Dictionary<int, List<int>> _matchIndexes;
        private HashSet<int> _rangesWithFormulas;

        private Queue<int> _flattenedRangesOrder;
        private Queue<int> _matchIndexesOrder;

        private Dictionary<string, int> _keyToId;
        private int _nextId = 1;

        /// <summary>
        /// Creates a new RangeCriteriaCache with specified limits
        /// </summary>
        /// <param name="package"></param>The ExcelPackage where the cache operates.</param>
        /// <param name="maxFlattenedRanges">Maximum number of flattened ranges to cache (default 100)</param>
        /// <param name="maxMatchIndexes">Maximum number of matchIndexes to cache (default 1000)</param>
        public RangeCriteriaCache(ExcelPackage package, int maxFlattenedRanges = DEFAULT_MAX_FLATTENED_RANGES,
                                   int maxMatchIndexes = DEFAULT_MAX_MATCH_INDEXES)
        {
            _package = package;
            _maxFlattenedRanges = maxFlattenedRanges;
            _maxMatchIndexes = maxMatchIndexes;
        }

        /// <summary>
        /// Gets a flattened range from cache, or null if not cached
        /// </summary>
        public List<object> GetFlattenedRange(FormulaRangeAddress address)
        {
            if (_flattenedRanges == null || address == null) return null;

            var key = CreateRangeKey(address);
            if (_keyToId != null && _keyToId.TryGetValue(key, out var id))
            {
                if (_flattenedRanges.TryGetValue(id, out var cached))
                {
                    return cached;
                }
            }
            return null;
        }

        /// <summary>
        /// Stores a flattened range in cache with FIFO eviction if cache is full
        /// </summary>
        public void SetFlattenedRange(FormulaRangeAddress address, List<object> flattenedRange)
        {
            if (address == null || flattenedRange == null) return;

            if (_flattenedRanges == null)
            {
                _flattenedRanges = new Dictionary<int, List<object>>();
                _flattenedRangesOrder = new Queue<int>();
                _keyToId = new Dictionary<string, int>();
            }

            var key = CreateRangeKey(address);

            // Get or create ID for this key
            if (!_keyToId.TryGetValue(key, out var id))
            {
                id = _nextId++;
                _keyToId[key] = id;

                // Check if cache is full and evict oldest entry (FIFO)
                if (_flattenedRanges.Count >= _maxFlattenedRanges)
                {
                    var oldestId = _flattenedRangesOrder.Dequeue();
                    _flattenedRanges.Remove(oldestId);
                    // Note: we keep the key in _keyToId to avoid ID reuse within same calculation
                }

                _flattenedRangesOrder.Enqueue(id);
            }

            _flattenedRanges[id] = flattenedRange;
        }

        /// <summary>
        /// Gets cached match indexes for a range+criteria combination.
        /// Returns null if the range contains formulas (which might change during calculation).
        /// </summary>
        public List<int> GetMatchIndexes(FormulaRangeAddress rangeAddress, object criteriaValue)
        {
            if (_matchIndexes == null || rangeAddress == null || criteriaValue == null) return null;

            var rangeKey = CreateRangeKey(rangeAddress);
            if (_keyToId != null && _keyToId.TryGetValue(rangeKey, out var rangeId))
            {
                // Don't use cache if this range contains formulas that might change
                if (_rangesWithFormulas != null && _rangesWithFormulas.Contains(rangeId))
                {
                    return null;
                }
            }

            var key = CreateMatchIndexKey(rangeAddress, criteriaValue);
            if (_keyToId != null && _keyToId.TryGetValue(key, out var id))
            {
                if (_matchIndexes.TryGetValue(id, out var cached))
                {
                    return cached;
                }
            }
            return null;
        }

        /// <summary>
        /// Stores match indexes for a range+criteria combination with FIFO eviction if cache is full
        /// </summary>
        public void SetMatchIndexes(FormulaRangeAddress rangeAddress, object criteriaValue, List<int> matchIndexes)
        {
            if (rangeAddress == null || criteriaValue == null || matchIndexes == null) return;

            if (_keyToId == null)
            {
                _keyToId = new Dictionary<string, int>();
            }

            // Track if this range has formulas
            var rangeKey = CreateRangeKey(rangeAddress);
            if (rangeAddress.HasFormulas(_package))
            {
                if (!_keyToId.TryGetValue(rangeKey, out var rangeId))
                {
                    rangeId = _nextId++;
                    _keyToId[rangeKey] = rangeId;
                }

                if (_rangesWithFormulas == null)
                {
                    _rangesWithFormulas = new HashSet<int>();
                }
                _rangesWithFormulas.Add(rangeId);
                // Don't cache matchIndexes for ranges with formulas
                return;
            }

            if (_matchIndexes == null)
            {
                _matchIndexes = new Dictionary<int, List<int>>();
                _matchIndexesOrder = new Queue<int>();
            }

            var key = CreateMatchIndexKey(rangeAddress, criteriaValue);

            // Get or create ID for this key
            if (!_keyToId.TryGetValue(key, out var id))
            {
                id = _nextId++;
                _keyToId[key] = id;

                // Check if cache is full and evict oldest entry (FIFO)
                if (_matchIndexes.Count >= _maxMatchIndexes)
                {
                    var oldestId = _matchIndexesOrder.Dequeue();
                    _matchIndexes.Remove(oldestId);
                    // Note: we keep the key in _keyToId to avoid ID reuse within same calculation
                }

                _matchIndexesOrder.Enqueue(id);
            }

            _matchIndexes[id] = matchIndexes;
        }

        /// <summary>
        /// Clears all cached data. Should be called after calculation completes.
        /// </summary>
        public void Clear()
        {
            _flattenedRanges?.Clear();
            _matchIndexes?.Clear();
            _rangesWithFormulas?.Clear();
            _flattenedRangesOrder?.Clear();
            _matchIndexesOrder?.Clear();
            _keyToId?.Clear();
            _nextId = 1;
        }

        private string CreateRangeKey(FormulaRangeAddress address)
        {
            // Create a unique key for the range address
            return $"{address.WorksheetIx}:{address.FromRow}:{address.FromCol}:{address.ToRow}:{address.ToCol}";
        }

        private string CreateMatchIndexKey(FormulaRangeAddress rangeAddress, object criteriaValue)
        {
            // Create a unique key combining range address and criteria value
            var rangeKey = CreateRangeKey(rangeAddress);
            var criteriaKey = criteriaValue?.ToString() ?? "null";
            return $"{rangeKey}|{criteriaKey}";
        }
    }
}