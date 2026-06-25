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
using EPPlus.Fonts.OpenType.Integration;
using System.Collections.Generic;

namespace EPPlus.Export.Pdf.Layout
{
    internal class PdfCellBase
    {
        public bool Hidden;
        public PdfCellAlignmentData ContentAligmnet;
        public bool IsPrintTitleRow;
        public bool IsPrintTitleCol;

        public string Name { get; set; }
        public List<TextFragment> TextFragments { get; set; }
        public List<PdfShapedText> ShapedTexts { get; set; }
        public TextLineCollection TextLines { get; set; }
        public string Text { get; set; }
        public double TotalTextLength { get; set; }
        public double ColumnWidth { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public TextLayoutEngine TextLayoutEngine { get; set; }
    }
}
