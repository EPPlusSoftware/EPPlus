using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DynamicArray.LookupUtils
{
    internal class XlookupScanner
    {
        public XlookupScanner(object lookupValue, IRangeInfo lookupRange, LookupSearchMode searchMode, LookupMatchMode matchMode)
        {
            _lookupValue = lookupValue;
            _lookupRange = lookupRange;
            _searchMode = searchMode;
            _matchMode = matchMode;
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
            if (_searchMode != LookupSearchMode.StartingAtFirst && _searchMode != LookupSearchMode.ReverseStartingAtLast)
            {
                return -1;
            }
            return FindIndexInternal();
        }

        private int FindIndexInternal()
        {
            var direction = GetLookupDirection();
            int dimensionRows, maxItems;
            if (direction == LookupRangeDirection.Vertical)
            {
                dimensionRows = _lookupRange.Worksheet.Dimension.Rows;
                maxItems = _lookupRange.Size.NumberOfRows > dimensionRows ? dimensionRows : _lookupRange.Size.NumberOfRows;
            }
            else
            {
                dimensionRows = _lookupRange.Worksheet.Dimension.Columns;
                maxItems = _lookupRange.Size.NumberOfCols > dimensionRows ? dimensionRows : _lookupRange.Size.NumberOfCols;
            }
            int closestBelowIx = -1;
            int closestAboveIx = -1;
            object closestBelow = null;
            object closestAbove = null;
            var ix = _searchMode == LookupSearchMode.ReverseStartingAtLast ? maxItems - 1 : 0;

            while (ix >= 0)
            {
                object value = direction == LookupRangeDirection.Vertical ?
                    _lookupRange.GetOffset(ix, 0) :
                    _lookupRange.GetOffset(0, ix);
                var cr = _comparer.Compare(_lookupValue, value);
                if (cr == 0)
                {
                    return ix;
                }
                else if (cr < 0)
                {
                    if (closestAbove == null || _comparer.Compare(closestAbove, value) > 0)
                    {
                        closestAbove = value;
                        closestAboveIx = ix;
                    }
                }
                else
                {
                    if (closestBelow == null || _comparer.Compare(closestBelow, value) < 0)
                    {
                        closestBelow = value;
                        closestBelowIx = ix;
                    }
                }
                if (_searchMode == LookupSearchMode.StartingAtFirst)
                {
                    ix++;
                    if (ix >= maxItems)
                    {
                        ix = -1;
                    }
                }
                else
                {
                    ix--;
                }
            }
            if (_matchMode == LookupMatchMode.ExactMatchReturnNextLarger)
            {
                return closestAboveIx;
            }
            else if (_matchMode == LookupMatchMode.ExactMatchReturnNextSmaller)
            {
                return closestBelowIx;
            }
            return -1;
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
