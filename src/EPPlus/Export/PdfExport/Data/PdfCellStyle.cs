/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.Export.PdfExport.Data
{
    /// <summary>
    /// Holds the styles for the cell.
    /// xf styles is the base style and dxf styles is an override(sort of) that can come from tables or conditional formatting.
    /// They are used in the following priority:
    /// 1. xf style that is not default
    /// 2. dxf style if it exsist
    /// 3. deault xf style
    /// </summary>
    internal class PdfCellStyle
    {
        //Fill
        internal ExcelFill xfFill { get; set; }
        internal ExcelDxfFill dxfFill { get; set; }
        internal ExcelDxfFontBase dxfFontOverride { get; set; }

        //Borders
        internal ExcelBorderItem xfTop { get; set; }
        internal ExcelBorderItem xfBottom { get; set; }
        internal ExcelBorderItem xfLeft { get; set; }
        internal ExcelBorderItem xfRight { get; set; }
        internal bool DiagonalUp { get; set; }
        internal bool DiagonalDown { get; set; }
        internal ExcelBorderItem Diagonal { get; set; }
        internal ExcelDxfBorderItem dxfTop { get; set; }
        internal ExcelDxfBorderItem dxfBottom { get; set; }
        internal ExcelDxfBorderItem dxfLeft { get; set; }
        internal ExcelDxfBorderItem dxfRight { get; set; }
        internal ExcelDxfBorderItem dxfHorizontal { get; set; }
        internal ExcelDxfBorderItem dxfVertical { get; set; }
        internal bool SuppressTop { get; set; }
        internal bool SuppressBottom { get; set; }
        internal bool SuppressLeft { get; set; }
        internal bool SuppressRight { get; set; }

        //Fonts
        internal ExcelFont xfFont { get; set; }
        internal ExcelDxfFontBase dxfFont { get; set; }
    }
}
