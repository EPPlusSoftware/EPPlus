/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  6/4/2022         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class HtmlImage
    {
        public int WorksheetId { get; set; }
        public ExcelPicture Picture { get; set; }
        public int FromRow { get; set; }
        public int FromRowOff { get; set; }
        public int ToRow { get; set; }
        public int ToRowOff { get; set; }
        public int FromColumn { get; set; }
        public int FromColumnOff { get; set; }
        public int ToColumn { get; set; }
        public int ToColumnOff { get; set; }
    }
}
