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
using System.ComponentModel;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A collection of pivot table fields
    /// </summary>
    public class ExcelPivotTableFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
    {
        private readonly ExcelPivotTable _table;
        internal ExcelPivotTableFieldCollection(ExcelPivotTable table) :
            base()
        {
            _table = table;
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

        internal ExcelPivotTableField AddDateGroupField(eDateGroupBy groupBy, int index)
        {
            //Pivot field
            var topNode = _table.PivotTableXml.SelectSingleNode("//d:pivotFields", _table.NameSpaceManager);
            var fieldNode = _table.PivotTableXml.CreateElement("pivotField", ExcelPackage.schemaMain);
            fieldNode.SetAttribute("compact", "0");
            fieldNode.SetAttribute("outline", "0");
            fieldNode.SetAttribute("showAll", "0");
            fieldNode.SetAttribute("defaultSubtotal", "0");
            topNode.AppendChild(fieldNode);

            var field = new ExcelPivotTableField(_table.NameSpaceManager, fieldNode, _table, _table.Fields.Count, index);
            field.DateGrouping = groupBy;
            _list.Add(field);
            return field;
        }

        internal void AddDateGroupField(eDateGroupBy dateGrouping, object baseIndex)
        {
            throw new NotImplementedException();
        }
    }
}