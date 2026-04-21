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
    internal class GroupLevel
    {
        public object Key { get; set; }
        public List<GroupLevel> Children { get; set; } = new List<GroupLevel>();
        public Dictionary<string, GroupLevel> ChildDict { get; set; } = null;
        public List<string> ChildOrder { get; set; } = null;
        public List<GroupRow> Rows { get; set; } = new List<GroupRow>();
        public object SubtotalValue { get; set; }
        public List<object[]> SubtotalValues { get; set; } = new List<object[]>(); // [function][valueCol]
        public bool IsLeaf => Children.Count == 0;
    }
}
