/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  19/3/2026         EPPlus Software AB           EPPlus v8.6
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.GroupingFunctions
{
    internal class GroupRow
    {
        public object[] KeyParts { get; set; }
        public List<object[]> Values { get; set; } = new List<object[]>();
        public object AggregatedValue { get; set; }
        public List<object[]> AggregatedValues { get; set; } = new List<object[]>(); // [function][valueCol]
    }
}
