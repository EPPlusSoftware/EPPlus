/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/10/2026         EPPlus Software AB       Initial release
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;

namespace OfficeOpenXml
{
    /// <summary>
    /// Settings for formula calculation caching and optimization
    /// </summary>
    public class ExcelCalculationCacheSettings
    {
        /// <summary>
        /// Maximum number of flattened ranges to cache for RangeCriteria functions (SUMIFS, AVERAGEIFS, COUNTIFS).
        /// Default is 100. Set to 0 to disable flattened range caching.
        /// </summary>
        public int MaxFlattenedRanges { get; set; } = RangeCriteriaCache.DEFAULT_MAX_FLATTENED_RANGES;

        /// <summary>
        /// Maximum number of match indexes to cache for RangeCriteria functions (SUMIFS, AVERAGEIFS, COUNTIFS).
        /// Default is 1000. Set to 0 to disable match index caching.
        /// </summary>
        public int MaxMatchIndexes { get; set; } = RangeCriteriaCache.DEFAULT_MAX_MATCH_INDEXES;

        /// <summary>
        /// Gets the default configuration for Excel calculation cache settings.
        /// </summary>
        /// <remarks>Use this property to obtain a standard set of cache settings suitable for most
        /// scenarios. The returned instance is immutable and can be shared safely across multiple components.</remarks>
        public static ExcelCalculationCacheSettings Default { get; } = new ExcelCalculationCacheSettings();

    }
}
