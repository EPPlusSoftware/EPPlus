/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// Defines if a table style applies to a Table / PivotTable or Both
    /// </summary>
    public enum eTableNamedStyleAppliesTo
    {
        /// <summary>
        /// The named style applies to tables only
        /// </summary>
        Tables,
        /// <summary>
        /// The named style applies to pivot tables only
        /// </summary>
        PivotTables,
        /// <summary>
        /// The named style can be applied to both tables and pivot tables
        /// </summary>
        TablesAndPivotTables
    }
}