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
namespace OfficeOpenXml.Export.PdfExport.Data
{
    internal class PdfDataBar
    {
        internal double Percentage { get; set; }
        internal bool IsNegative { get; set; }
        internal double AxisFraction { get; set; }
        internal System.Drawing.Color FillColor { get; set; }
        internal System.Drawing.Color? BorderColor { get; set; }
        internal System.Drawing.Color? AxisColor { get; set; }
        internal bool Gradient { get; set; }
        internal bool IsRightToLeft { get; set; }
    }
}