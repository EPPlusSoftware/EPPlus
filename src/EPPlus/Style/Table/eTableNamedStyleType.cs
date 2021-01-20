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
    /// The type of custom named style for tables and pivot tables
    /// </summary>
    public enum eTableNamedStyleType
    {
        /// <summary>
        /// A custom named style for tables
        /// </summary>
        Table,
        /// <summary>
        /// A custom named style for  pivot tables
        /// </summary>
        PivotTable,
        /// <summary>
        /// A custom named style for tables and  pivot tables
        /// </summary>
        PivotTableAndTable
    }
}