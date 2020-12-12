using System;
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class ColumnInfo
    {
        public int SortOrder { get; set; }
        public int Index { get; set; }

        public MemberInfo MemberInfo { get; set; }

        public string Formula { get; set; }

        public string FormulaR1C1 { get; set; }

        public string Header { get; set; }

        public string NumberFormat { get; set; }
    }
}
