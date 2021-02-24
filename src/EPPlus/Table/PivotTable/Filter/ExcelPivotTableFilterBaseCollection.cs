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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable.Filter
{
    public abstract class ExcelPivotTableFilterBaseCollection : IEnumerable<ExcelPivotTableFilter>
    {
        protected internal List<ExcelPivotTableFilter> _filters = new List<ExcelPivotTableFilter>();
        protected internal readonly ExcelPivotTable _table;
        protected internal readonly ExcelPivotTableField _field;
        internal ExcelPivotTableFilterBaseCollection(ExcelPivotTable table)
        {
            _table = table;
            var filtersNode = _table.GetNode("d:filters");
            if (filtersNode != null)
            {
                foreach (XmlNode node in filtersNode.ChildNodes)
                {
                    var f =new ExcelPivotTableFilter(_table.NameSpaceManager, node, _table.WorkSheet.Workbook.Date1904);
                    table.SetNewFilterId(f.Id);
                    _filters.Add(f);
                }
            }
        }
        internal ExcelPivotTableFilterBaseCollection(ExcelPivotTableField field)
        {            
            _field = field;
            _table = field._pivotTable;

            foreach(var filter in _table.Filters)
            {
                if(filter.Fld==field.Index)
                {
                    _filters.Add(filter);
                }
            }
        }
        public IEnumerator<ExcelPivotTableFilter> GetEnumerator()
        {
            return _filters.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _filters.GetEnumerator();
        }

        internal XmlNode GetOrCreateFiltersNode()
        {
            return _table.CreateNode("d:filters");
        }
        internal protected ExcelPivotTableFilter CreateFilter()
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

        public int Count 
        { 
            get
            {
                return _filters.Count;
            }
        }
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
