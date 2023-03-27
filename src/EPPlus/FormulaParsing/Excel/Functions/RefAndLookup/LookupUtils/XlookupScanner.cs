using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils
{
    internal class XlookupScanner
    {
        public XlookupScanner(object lookupValue, IRangeInfo lookupRange, LookupSearchMode searchMode, LookupMatchMode matchMode, LookupRangeDirection direction)
        {
            _lookupValue= lookupValue;
            _lookupRange= lookupRange;
            _searchMode= searchMode;
            _matchMode= matchMode;
            _comparer = new LookupComparer(matchMode);
        }

        private readonly object _lookupValue;
        private readonly IRangeInfo _lookupRange;
        private readonly LookupSearchMode _searchMode;
        private readonly LookupMatchMode _matchMode;
        private readonly IComparer<object> _comparer;

        private LookupRangeDirection GetLookupDirection()
        {
            var result = LookupRangeDirection.Vertical;
            if (_lookupRange.Size.NumberOfCols > 1)
            {
                result = LookupRangeDirection.Horizontal;
            }
            return result;
        }

        public int FindIndex()
        {
            var direction = GetLookupDirection();
            if(direction == LookupRangeDirection.Vertical)
            {
                return FindVertical();
            }
            else
            {
                return FindHorizontal();
            }
        }

        private int FindVertical()
        {
            var dimensionRows = _lookupRange.Worksheet.Dimension.Rows;
            var maxRows = _lookupRange.Size.NumberOfRows > dimensionRows ? dimensionRows : _lookupRange.Size.NumberOfRows;
            int closestBelowIx = -1;
            int closestAboveIx = -1;
            object closestBelow = null;
            object closestAbove = null;

            for (var rowIx = 0; rowIx < maxRows; rowIx++)
            {
                var value = _lookupRange.GetOffset(rowIx, 0);
                var cr = _comparer.Compare(_lookupValue, value);
                if(cr == 0)
                {
                    return rowIx;
                }
                else if(cr < 0)
                {
                    if(closestBelow == null || _comparer.Compare(closestBelow, value) < 0)
                    {
                        closestBelow = value;
                        closestBelowIx = rowIx;
                    }
                    if(closestAbove == null || _comparer.Compare(closestAbove, value) < 0)
                    {
                        closestAbove = value;
                        closestAboveIx = rowIx;
                    }
                }
            }
            if (_matchMode == LookupMatchMode.ExactMatchReturnNextLarger)
            {
                return closestAboveIx;
            }  
            else
            {
                return closestBelowIx;
            }
        }

        private int FindHorizontal()
        {
            var dimensionCols = _lookupRange.Worksheet.Dimension.Columns;
            var maxCols = _lookupRange.Size.NumberOfCols > dimensionCols ? dimensionCols : _lookupRange.Size.NumberOfCols;
            int closestBelowIx = -1;
            int closestAboveIx = -1;
            object closestBelow = null;
            object closestAbove = null;

            for (var colIx = 0; colIx < maxCols; colIx++)
            {
                var value = _lookupRange.GetOffset(0, colIx);
                var cr = _comparer.Compare(_lookupValue, value);
                if (cr == 0)
                {
                    return colIx;
                }
                else if (cr < 0)
                {
                    if (closestBelow == null || _comparer.Compare(closestBelow, value) < 0)
                    {
                        closestBelow = value;
                        closestBelowIx = colIx;
                    }
                    if (closestAbove == null || _comparer.Compare(closestAbove, value) < 0)
                    {
                        closestAbove = value;
                        closestAboveIx = colIx;
                    }
                }
            }
            if (_matchMode == LookupMatchMode.ExactMatchReturnNextLarger)
            {
                return closestAboveIx;
            }
            else
            {
                return closestBelowIx;
            }
        }
    }
}
