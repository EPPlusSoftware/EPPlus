using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class SlicerPivotTableCollection : IEnumerable<ExcelPivotTable>
    {
        internal List<ExcelPivotTable> _list = new List<ExcelPivotTable>();
        public IEnumerator<ExcelPivotTable> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        public ExcelPivotTable this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        public void Add(ExcelPivotTable pivotTable)
        {
            if(_list.Count>0)
            {
                if(pivotTable.CacheDefinition.CacheSource!=_list[0].CacheDefinition.CacheSource)
                {
                    throw (new InvalidOperationException("Pivot tables sharing the same slicer must have the same source data cache."));
                }
            }
            _list.Add(pivotTable);
        }
        public void Remove(ExcelPivotTable pivotTable)
        {
            _list.Remove(pivotTable);
        }
        public void RemoveAt(int index)
        {
            _list.RemoveAt(index);
        }
    }
}
