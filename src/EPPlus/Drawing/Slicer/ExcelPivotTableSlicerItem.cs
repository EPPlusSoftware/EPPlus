using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Slicer
{
    /// <summary>
    /// Represents a pivot table slicer item.
    /// </summary>
    public class ExcelPivotTableSlicerItem
    {
        private ExcelPivotTableSlicerCache _cache;
        private int _index;

        internal ExcelPivotTableSlicerItem(ExcelPivotTableSlicerCache cache, int index)
        {
            _cache = cache;
            _index = index;
        }
        /// <summary>
        /// The value of the item
        /// </summary>
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
        /// <summary>
        /// If the value is hidden 
        /// </summary>
        public bool Hidden 
        { 
            get
            {
                if (_index >= _cache._field.Items.Count)
                {
                    throw(new IndexOutOfRangeException());
                }
                var ix = _cache._field.Items.CacheLookup[_index].First();
                return _cache._field.Items[ix].Hidden;
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
                    var ix = fld.Items.CacheLookup[_index].First();
                    fld.Items[ix].Hidden = value;
                }
            }
        }
    }
}
