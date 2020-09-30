using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelPivotTableSlicerItem
    {
        private ExcelPivotTableSlicerCache _cache;
        private int _index;

        public ExcelPivotTableSlicerItem(ExcelPivotTableSlicerCache cache, int index)
        {
            _cache = cache;
            _index = index;
        }
        public object Value 
        { 
            get
            {
                if (_index >= _cache._field.Items.Count)
                {
                    return null;
                }
                return _cache._field.Items[_index].Value;
            }
        }
        public bool Hidden 
        { 
            get
            {
                if (_index >= _cache._field.Items.Count)
                {
                    throw(new IndexOutOfRangeException());
                }
                return _cache._field.Items[_index].Hidden;
            }
            set
            {
                if (_index >= _cache.Data.Items.Count)
                {
                    throw (new IndexOutOfRangeException());
                }
                foreach (var pt in _cache.PivotTables)
                {
                    var fld = pt.Fields[_cache._field.Index];
                    if (_index >= fld.Items.Count || fld.Items[_index].Type != Table.PivotTable.eItemType.Data) continue;
                    fld.Items[_index].Hidden = value;
                }
            }
        }
    }
}
