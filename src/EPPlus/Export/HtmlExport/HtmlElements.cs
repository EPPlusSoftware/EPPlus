/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal static class HtmlElements
    {
        public const string Body = "body";

        public const string Table = "table";
        public const string Thead = "thead";
        public const string TFoot = "tfoot";
        public const string Tbody = "tbody";
        public const string TableRow = "tr";
        public const string TableHeader = "th";
        public const string TableData = "td";
        public const string A = "a";
        public const string Span = "span";
    }
}
