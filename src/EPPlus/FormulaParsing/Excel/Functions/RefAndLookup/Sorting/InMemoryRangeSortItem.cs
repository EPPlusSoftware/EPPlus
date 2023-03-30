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
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Sorting
{
    [DebuggerDisplay("oix: {OriginalIndex}, v: {Value}")]
    internal class InMemoryRangeSortItem
    {
        public InMemoryRangeSortItem(object value, int originalIndex)
        {
            Value = value;
            OriginalIndex = originalIndex;
        }
    
        public object Value { get; set; }

        public int OriginalIndex { get; set; }
    }
}
