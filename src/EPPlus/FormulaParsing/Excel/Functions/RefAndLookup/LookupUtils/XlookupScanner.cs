/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/3/2023         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static OfficeOpenXml.FormulaParsing.Excel.Functions.Math.RoundingHelper;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils
{
    internal class XlookupScanner
    {
        public XlookupScanner(
            object lookupValue,
            IRangeInfo lookupRange,
            LookupSearchMode searchMode,
            LookupMatchMode matchMode)
        {
            _lookupValue = lookupValue;
            _lookupRange = lookupRange;
            _searchMode = searchMode;
            _matchMode = matchMode;
            _comparer = new LookupComparer(matchMode);
            _direction = LookupRangeDirection.Undefined;
        }

        public XlookupScanner(
            object lookupValue, 
            IRangeInfo lookupRange, 
            LookupSearchMode searchMode, 
            LookupMatchMode matchMode,
            LookupRangeDirection direction)
        {
            _lookupValue = lookupValue;
            _lookupRange = lookupRange;
            _searchMode = searchMode;
            _matchMode = matchMode;
            _direction = direction;
            _comparer = new LookupComparer(matchMode);
        }

        private readonly object _lookupValue;
        private readonly IRangeInfo _lookupRange;
        private readonly LookupSearchMode _searchMode;
        private readonly LookupMatchMode _matchMode;
        private readonly LookupRangeDirection _direction;
        private readonly IComparer<object> _comparer;

        private LookupRangeDirection GetLookupDirection()
        {
            if(_direction != LookupRangeDirection.Undefined)
            {
                return _direction;
            }
            if (_lookupRange.Size.NumberOfCols > 1)
            {
                return LookupRangeDirection.Horizontal;
            }
            return LookupRangeDirection.Vertical;
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
            int  maxItems;
            if (direction == LookupRangeDirection.Vertical)
            {
                maxItems = GetMaxItemsRow(_lookupRange);
            }
            else
            {
                //dimensionItems = _lookupRange.Dimension.ToCol - _lookupRange.Dimension.FromCol + 1;
                //maxItems = _lookupRange.Size.NumberOfCols > dimensionItems ? dimensionItems : _lookupRange.Size.NumberOfCols;
                maxItems = GetMaxItemsColumns(_lookupRange);
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

        private int GetMaxItemsRow(IRangeInfo lookupRange)
        {
            if (lookupRange.Address.ToRow > lookupRange.Dimension.ToRow)
            {
                return lookupRange.Dimension.ToRow - lookupRange.Address.FromRow + 1;
            }            
            return _lookupRange.Size.NumberOfRows;
        }
        private int GetMaxItemsColumns(IRangeInfo lookupRange)
        {
            if (lookupRange.Address.ToCol > lookupRange.Dimension.ToCol)
            {
                return lookupRange.Dimension.ToCol - lookupRange.Address.FromCol + 1;
            }
            return _lookupRange.Size.NumberOfCols;
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
