/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2020         EPPlus Software AB       EPPlus 5.4
 *************************************************************************************************/
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable.Filter
{
    /// <summary>
    /// The base collection for pivot table filters
    /// </summary>
    public abstract class ExcelPivotTableFilterBaseCollection : IEnumerable<ExcelPivotTableFilter>
    {
        internal List<ExcelPivotTableFilter> _filters;
        internal readonly ExcelPivotTable _table;
        internal readonly ExcelPivotTableField _field;
        internal ExcelPivotTableFilterBaseCollection(ExcelPivotTable table)
        {
            _table = table;
            ReloadTable();
        }
        internal ExcelPivotTableFilterBaseCollection(ExcelPivotTableField field)
        {
            _filters = new List<ExcelPivotTableFilter>();
            _field = field;
            _table = field.PivotTable;

            foreach(var filter in _table.Filters)
            {
                if(filter.Fld==field.Index)
                {
                    _filters.Add(filter);
                }
            }
        }
        /// <summary>
        /// Reloads the collection from the xml.
        /// </summary>
        internal void ReloadTable()
        {
            _filters = new List<ExcelPivotTableFilter>();
            var filtersNode = _table.GetNode("d:filters");
            if (filtersNode != null)
            {
                foreach (XmlNode node in filtersNode.ChildNodes)
                {
                    var f = new ExcelPivotTableFilter(_table.NameSpaceManager, node, _table.WorkSheet.Workbook.Date1904);
                    _table.SetNewFilterId(f.Id);
                    _filters.Add(f);
                }
            }
        }
        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator that can be used to iterate through the collection.</returns>
        public IEnumerator<ExcelPivotTableFilter> GetEnumerator()
        {
            return _filters.GetEnumerator();
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>An enumerator that can be used to iterate through the collection.</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _filters.GetEnumerator();
        }

        internal XmlNode GetOrCreateFiltersNode()
        {
            return _table.CreateNode("d:filters");
        }
        internal ExcelPivotTableFilter CreateFilter()
        {
            var topNode = GetOrCreateFiltersNode();
            var filterNode = topNode.OwnerDocument.CreateElement("filter", ExcelPackage.schemaMain);
            topNode.AppendChild(filterNode);
            var filter = new ExcelPivotTableFilter(_field.NameSpaceManager, filterNode, _table.WorkSheet.Workbook.Date1904)
            {
                EvalOrder = -1,
                Fld = _field.Index,
                Id = _table.GetNewFilterId()
            };
            return filter;
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count 
        { 
            get
            {
                return _filters.Count;
            }
        }
        /// <summary>
        /// The indexer for the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns></returns>
        public ExcelPivotTableFilter this[int index]
        {
            get
            {
                if (index < 0 || index >= _filters.Count)
                    throw (new ArgumentOutOfRangeException());
                
                return _filters[index];
            }
        }
    }
}
