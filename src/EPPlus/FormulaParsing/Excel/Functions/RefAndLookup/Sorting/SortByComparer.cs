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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting
{
    internal class SortByComparer : LookupComparerBase
    {
        public SortByComparer() : base(LookupMatchMode.ExactMatch)
        {
        }

        public override int Compare(object x, object y)
        {
            return Compare(x, y, 1);
        }

        public override int Compare(object x, object y, int sortOrder)
        {
            // null values should always be sorted last
            if(x == null && y != null)
            {
                return 1;
            }
            else if(x != null && y == null)
            {
                return -1;
            }
            return CompareObjects(x, y) * sortOrder;
        }
    }
}
