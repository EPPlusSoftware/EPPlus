using System;
using System.Collections.Generic;
using System.Text;
#if !NET35
using System.Collections.Concurrent;
#endif

namespace OfficeOpenXml.Utils.AutofitCols
{
    internal class ColumnWidthMap
    {
        private readonly object _syncRoot = new object();
#if NET35
        private Dictionary<int, double> _maxWidths = new Dictionary<int, double>();
        
#else
        private ConcurrentDictionary<int, double> _maxWidths = new ConcurrentDictionary<int, double>();
#endif

        internal IDictionary<int, double> GetResult()
        {
            return _maxWidths;
        }

#if NET35
        internal void AddMeasurement(int col, double width)
        {
            if(!_maxWidths.ContainsKey(col))
            {
                lock(_syncRoot)
                {
                    _maxWidths[col] = width;
                }
            }
            else
            {
                if(width > _maxWidths[col])
                {
                    lock(_syncRoot)
                    {
                        _maxWidths[col] = width;
                    }
                }
            }
        }
#else
        internal void AddMeasurement(int col, double width)
        {
            if(!_maxWidths.ContainsKey(col))
            {
                _maxWidths[col] = width;
            }
            else if(width > _maxWidths[col])
            {
                _maxWidths[col] = width;
            }
        }
#endif
    }
}
