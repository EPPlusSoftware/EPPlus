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
using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Fonts.OpenType.Integration;
using OfficeOpenXml.Export.PdfExport.TextShaping;
using OfficeOpenXml.Style;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.PdfExport.Data
{
    internal class PdfCell
    {
        public string Name { get; set; }
        public bool Hidden;
        public PdfCellStyle CellStyle;
        public PdfCellAlignmentData ContentAligmnet;
        public List<TextFragment> TextFragments { get; set; }
        public List<PdfShapedText> ShapedTexts { get; set; }
        public TextLineCollection TextLines { get; set; }
        public string Text { get; set; }


        public double TotalTextLength { get; set; }
        public double ColumnWidth { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }

        public TextLayoutEngine TextLayoutEngine { get; set; }

        public bool Merged;
        public PdfCell Main;
        public ExcelAddressBase MergedAddress;

        public bool IsPrintTitleRow;
        public bool IsPrintTitleCol;
    }

    internal class PdfCellAlignmentData
    {
        public ExcelHorizontalAlignment HorizontalAlignment = ExcelHorizontalAlignment.General;
        public ExcelVerticalAlignment VerticalAlignment = ExcelVerticalAlignment.Bottom;
        public int Indent = 0;
        public bool WrapText = false;
        public bool ShrinkToFit = false;
        public int TextRotation = 0;
        public ExcelReadingOrder TextDirection = ExcelReadingOrder.ContextDependent;
        public bool IsVertical = false;

        public PdfCellAlignmentData() { }
    }
}
