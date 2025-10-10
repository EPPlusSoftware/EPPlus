/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB       Initial release EPPlus 8.3
 *************************************************************************************************/
using System;

namespace OfficeOpenXml.Data.Connection
{
    public class ExcelConnectionOlapProperties
    {
        public bool Local { get; set; } = false;
        public string LocalConnection { get; set; }
        public bool LocalRefresh { get; set; } = true;
        public bool SendLocale { get; set; } = false;
        public int? RowDrillCount { get; set; }
        public bool ServerFill { get; set; } = true;
        public bool ServerNumberFormat { get; set; } = true;
        public bool ServerFont { get; set; } = true;
        public bool ServerFontColor { get; set; } = true;
        }
}