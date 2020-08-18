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
using System;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A collection of pivot table fields
    /// </summary>
    public class ExcelPivotTableFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
    {
        internal ExcelPivotTableFieldCollection(ExcelPivotTable table, string topNode) :
            base(table, 0)
        {

        }
        /// <summary>
        /// Indexer by name
        /// </summary>
        /// <param name="name">The name</param>
        /// <returns>The pivot table field</returns>
        public ExcelPivotTableField this[string name]
        {
            get
            {
                foreach (var field in _list)
                {
                    if (field.Name.Equals(name,StringComparison.OrdinalIgnoreCase))
                    {
                        return field;
                    }
                }
                return null;
            }
        }
        /// <summary>
        /// Returns the date group field.
        /// </summary>
        /// <param name="GroupBy">The type of grouping</param>
        /// <returns>The matching field. If none is found null is returned</returns>
        public ExcelPivotTableField GetDateGroupField(eDateGroupBy GroupBy)
        {
            foreach (var fld in _list)
            {
                if (fld.Grouping is ExcelPivotTableFieldDateGroup && (((ExcelPivotTableFieldDateGroup)fld.Grouping).GroupBy) == GroupBy)
                {
                    return fld;
                }
            }
            return null;
        }
        /// <summary>
        /// Returns the numeric group field.
        /// </summary>
        /// <returns>The matching field. If none is found null is returned</returns>
        public ExcelPivotTableField GetNumericGroupField()
        {
            foreach (var fld in _list)
            {
                if (fld.Grouping is ExcelPivotTableFieldNumericGroup)
                {
                    return fld;
                }
            }
            return null;
        }
    }
}