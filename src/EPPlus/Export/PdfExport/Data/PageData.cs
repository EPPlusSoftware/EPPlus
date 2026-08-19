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
using OfficeOpenXml.Export.PdfExport.TextMapping;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using EPPlus.Export.Pdf.Layout;
using OfficeOpenXml.Export.PdfExport.Layout;


namespace OfficeOpenXml.Export.PdfExport.Data
{
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
        public Dictionary<string, MergedCellDrawInfo> MergedCells;
        public List<PrintTitleCellDraw> PrintTitleCells;
        public List<GridLine> PrintTitleGridLines;
        public List<PrintTitleHeadingDraw> PrintTitleHeadings;
        public List<SpillCellDraw> SpillCells;
        public List<PrintTitleCellDraw> PrintTitleBorders;
        public double[] RowHeights;
        public double HeadingWidth;
        public double HeadingHeight;
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
        public List<PdfDrawing> Drawings;
        public int Count
        {
            get { return Width * Height; }
        }
    }
}
