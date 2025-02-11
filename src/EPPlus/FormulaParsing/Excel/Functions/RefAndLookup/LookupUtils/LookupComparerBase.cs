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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.LookupUtils
{
    internal abstract class LookupComparerBase : IComparer<object>
    {
        public LookupComparerBase(LookupMatchMode matchMode)
        {
            _matchMode = matchMode;
        }

        private readonly LookupMatchMode _matchMode;
        private readonly ValueMatcher _vm = new WildCardValueMatcher2();

        public abstract int Compare(object x, object y);

        public virtual int Compare(object x, object y, int sortOrder)
        {
            int v = 0;
            if (_matchMode == LookupMatchMode.Wildcard || _matchMode == LookupMatchMode.ExactMatchWithWildcard)
            {
                v = _vm.IsMatch(x, y);
            }
            else
            {
                v = ComparerUtil.CompareObjects(x, y);
            }
            return v * (sortOrder > -1 ? 1 : -1);
        }
    }
}
