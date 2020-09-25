using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelPivotTableSlicerItem
    {
        private ExcelPivotTableSlicer _slicer;
        private int _index;

        public ExcelPivotTableSlicerItem(ExcelPivotTableSlicer slicer, int index)
        {
            _slicer = slicer;
            _index = index;
        }
        public object Value 
        { 
            get
            {
                if (_index >= _slicer._field.Items.Count)
                {
                    return null;
                }
                return _slicer._field.Items[_index].Value;
            }
        }
        public bool Hidden 
        { 
            get
            {
                if (_index >= _slicer._field.Items.Count)
                {
                    throw(new IndexOutOfRangeException());
                }
                return _slicer._field.Items[_index].Hidden;
            }
            set
            {
                if (_index >= _slicer.Cache.Data.Items.Count)
                {
                    throw (new IndexOutOfRangeException());
                }
                foreach (var pt in _slicer.Cache.PivotTables)
                {
                    var fld = pt.Fields[_slicer._field.Index];
                    if (_index >= fld.Items.Count || fld.Items[_index].Type != Table.PivotTable.eItemType.Data) continue;
                    fld.Items[_index].Hidden = value;
                }
            }
        }
    }
}
