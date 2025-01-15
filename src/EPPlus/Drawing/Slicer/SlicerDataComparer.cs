/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  14/1/2025         EPPlus Software AB           EPPlus v7
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting
{
    internal class SlicerDataComparer : IComparer<ExcelPivotTableFieldItem>
    {
        eCrossFilter _crossFilter;
        public SlicerDataComparer(eCrossFilter crossFilter)
        {
            _crossFilter = crossFilter;
        }

        public int Compare(ExcelPivotTableFieldItem x, ExcelPivotTableFieldItem y)
        {
            return Compare(x, y, 1);
        }        
        public virtual int Compare(ExcelPivotTableFieldItem x, ExcelPivotTableFieldItem y, int sortOrder)
        {
            if (x.Value == null && y.Value != null && x.Hidden)
            {
                return 1;
            }
            else if (x.Value != null && y.Value == null)
            {
                return -1;
            }
            if (_crossFilter == eCrossFilter.ShowItemsWithDataAtTop)
            { 

            }
            return ComparerUtil.CompareObjects(x.Value, y.Value) * sortOrder;
        }
    }
}
