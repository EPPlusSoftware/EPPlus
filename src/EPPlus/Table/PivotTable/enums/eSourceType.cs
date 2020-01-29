/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Source type for a pivottable
    /// </summary>
    public enum eSourceType
    {
        /// <summary>
        /// The cache contains data that consolidates ranges
        /// </summary>
        Consolidation,
        /// <summary>
        /// The cache contains data from an external data source
        /// </summary>
        External,
        /// <summary>
        /// The cache contains a scenario summary report
        /// </summary>
        Scenario,
        /// <summary>
        /// The cache contains worksheet data
        /// </summary>
        Worksheet
    }
}