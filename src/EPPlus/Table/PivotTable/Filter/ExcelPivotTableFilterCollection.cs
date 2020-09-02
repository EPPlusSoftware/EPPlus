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
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Xml;

namespace EPPlusTest.Table.PivotTable.Filter
{
    public class ExcelPivotTableFilterCollection : IEnumerable<ExcelPivotTableFilter>
    {
        private readonly ExcelPivotTable _table;
        internal ExcelPivotTableFilterCollection(ExcelPivotTable table)
        {
            _table = table;
            var filtersNode = _table.GetNode("d:filters");
            if(filtersNode!=null)
            {
                foreach (XmlNode node in filtersNode.ChildNodes)
                {
                    _filters.Add(new ExcelPivotTableFilter(_table.NameSpaceManager, node));
                }
            }
        }
        private List<ExcelPivotTableFilter> _filters=new List<ExcelPivotTableFilter>();
        public IEnumerator<ExcelPivotTableFilter> GetEnumerator()
        {
            return _filters.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _filters.GetEnumerator();
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
