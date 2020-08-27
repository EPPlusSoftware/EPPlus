using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelPivotTableSlicerItemCollection : IEnumerable<ExcelPivotTableSlicerItem>
    {
        public readonly ExcelPivotTableSlicer _slicer;
        public ExcelPivotTableSlicerItemCollection(ExcelPivotTableSlicer slicer)
        {
            _slicer = slicer;
        }
        public IEnumerator<ExcelPivotTableSlicerItem> GetEnumerator()
        {
            throw new NotImplementedException();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
        public int Count 
        { 
            get
            {
                return _slicer._field.Items.Count;
            }
        }
        public ExcelPivotTableSlicerItem this[int index]
        {
            get
            {
                return new ExcelPivotTableSlicerItem(_slicer, index);
            }
        }
    }
}
    