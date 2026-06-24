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
using System.Drawing;
using EPPlus.Export.Pdf.Enums;

namespace EPPlus.Export.Pdf.Layout
{
    internal class PdfCellFillData
    {
        public string id;
        public Color BackgroundColor = Color.Empty;
        public ExcelFillStyle PatternStyle = ExcelFillStyle.None;
        public Color PatternColor = Color.Black;
        //Fill Effects
        public PdfCellGradientFillData GradientFillData = null;
        public bool enhanceGridLine = false;
        public PdfCellFillData() { }
    }
}
