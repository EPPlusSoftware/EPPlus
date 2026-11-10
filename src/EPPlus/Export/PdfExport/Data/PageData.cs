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
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Settings;
using OfficeOpenXml.Export.PdfExport.Layout;
using OfficeOpenXml.Export.PdfExport.TextMapping;
using OfficeOpenXml.Style;
using System.Collections.Generic;


namespace OfficeOpenXml.Export.PdfExport.Data
{
    internal struct PrintTitleCellDraw
    {
        public PdfCell Cell;
        public double X;
        public double Y;
        public double Width;
        public double Height;
        public double ClipX;
        public double ClipY;
        public double ClipWidth;
        public double ClipHeight;
    }

    internal struct PrintTitleHeadingDraw
    {
        public bool IsRow;
        public int Index;
        public double X;
        public double Y;
        public double Width;
        public double Height;
    }

    internal struct SpillCellDraw
    {
        public PdfCell Cell;
        public double X, Y, Width, Height;
        public double ClipX, ClipY, ClipWidth, ClipHeight;
        public bool IsPrintTitle;
    }

    internal struct Page
    {
        public int FromRow;
        public int FromColumn;
        public int ToRow;
        public int ToColumn;
        public bool HasPrintTitle;
        public double PrintTitleWidth;
        public double PrintTitleHeight;
        public PdfCellCollection Map;
        public PdfHeaderFooterCollection HeaderFooters;
        public PdfAdditionalContent AddedContent;
        public Dictionary<string, MergedCellDrawInfo> MergedCells;
        public List<ImageDrawInfo> Images;
        public List<PrintTitleCellDraw> PrintTitleCells;
        public List<GridLine> PrintTitleGridLines;
        public List<PrintTitleHeadingDraw> PrintTitleHeadings;
        public List<SpillCellDraw> SpillCells;
        public List<PrintTitleCellDraw> PrintTitleBorders;
        public double[] RowHeights;
        public double HeadingWidth;
        public double HeadingHeight;
        public double UsedWidth;
        public double UsedHeight;
        public bool IsCommentsPage;
    }

    internal struct Pages
    {
        public Page[] Page;
        public int Width;
        public int Height;
        public bool IsCommentsPage;
        public string HeadingFontName;
        public float HeadingFontSize;
        public ExcelFill HeadingFill;
        public PdfPageSettings Settings;
        public int SheetIndex;
        public string SheetName;
        public int Count
        {
            get { return Width * Height; }
        }
    }
}
