using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Drawing.Slicer
{
    public class ExcelPivotTableSlicerItemCollection : IEnumerable<ExcelPivotTableSlicerItem>
    {
        private readonly ExcelPivotTableSlicer _slicer;
        private readonly List<ExcelPivotTableSlicerItem> _items;
        internal ExcelPivotTableSlicerItemCollection(ExcelPivotTableSlicer slicer)
        {
            _slicer = slicer;
            _items = new List<ExcelPivotTableSlicerItem>();
            Refresh();
        }

        public void Refresh()
        {
            var cacheItems = _slicer._field.Cache.Grouping==null ? _slicer._field.Cache.SharedItems : _slicer._field.Cache.GroupItems;
            if(cacheItems.Count == _items.Count)
            {
                return;
            }
            else if(cacheItems.Count>_items.Count)
            {
                for (int i = _items.Count; i < cacheItems.Count; i++)
                {
                    _items.Add(new ExcelPivotTableSlicerItem(_slicer, i));
                }
            }
            else
            {                
                while(cacheItems.Count<_items.Count)
                {
                    _items.RemoveAt(_items.Count - 1);
                }
            }
        }

        public IEnumerator<ExcelPivotTableSlicerItem> GetEnumerator()
        {
            Refresh();
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            Refresh();
            return _items.GetEnumerator();
        }
        public int Count 
        { 
            get
            {
                return _items.Count;
            }
        }
        public ExcelPivotTableSlicerItem this[int index]
        {
            get
            {
                return _items[index];
            }
        }
        public ExcelPivotTableSlicerItem this[object value]
        {
            get
            {
                if(_slicer._field.Cache._cacheLookup.TryGetValue(value, out int ix))
                {
                    return _items[ix];
                }
                return null;
            }
        }
    }
}
    