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
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using EPPlus.Graphics.Units;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable.Filter;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfWorksheetLayout : Transform
    {
        public readonly double ZeroCharWidth;

        public PdfWorksheetLayout(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfDictionaries dictionaries)
        {
            double x = 0d, y = 0d, totalWidth = 0d;
            ZeroCharWidth = GetThemeFont0Width(worksheet);
            List<string> checkedMergedCells = new List<string>();
            for(int row = 1; row<= worksheet.Dimension._toRow; row++)
            {
                if(worksheet.Row(row).Hidden) continue;
                var height = UnitConversion.ExcelRowHeightToPoints(worksheet.Row(row).Height);
                x = 0d;
                for (int col = 1; col <= worksheet.Dimension._toCol; col++)
                {
                    if(worksheet.Column(col).Hidden) continue;
                    var width = UnitConversion.ExcelColumnWidthToPoints(worksheet.Column(col).Width, ZeroCharWidth);
                    var cell = worksheet.Cells[row, col];
                    var cellStyle = new PdfCellStyleOverride();
                    GetFillStyles(cell, cellStyle);
                    //CheckOverrideBorderStyle(cell, cellStyle);
                    PdfCellBorderLayout border = HandleEdgeBorders(cell, cellStyle, x, y, width, height);
                    if (cell.Merge)
                    {
                        //HandleMergedCell(worksheet, pageSettings, dictionaries, cell, checkedMergedCells, x, y);
                    }
                    else
                    {
                        HandleCell(pageSettings, dictionaries, cell, x, y, width, height, cellStyle);
                    }
                    //if (border != null) border.InitEdgeBorders(cell);
                    x += width;
                    totalWidth = System.Math.Max( x, totalWidth);
                }
                y -= height;
            }
            HandleDrawings(worksheet);
            Size = new Vector2(totalWidth, y);
        }

        private void CheckOverrideBorderStyle(ExcelRange cell, PdfCellStyleOverride tableStyle)
        {
            var tbl = cell.Worksheet.Tables.GetIntersectingRanges(cell);
            if (tbl.Count > 0)
            {
                var tblrng = tbl[0].Value.Range;
                var table = tbl[0].Value;

                int tblrow = 0;
                int tblcol = 0;
                tblrow = cell._fromRow - tblrng._fromRow;
                tblcol = cell._fromCol - tblrng._fromCol;

                if (tblrow == 0)
                {
                    tableStyle.borderStyleType |= TableBorderStyle.Top;
                }
                if (tblrng._toRow == cell._fromRow)
                {
                    tableStyle.borderStyleType |= TableBorderStyle.Bottom;
                }
                if (tblcol == 0)
                {
                    tableStyle.borderStyleType |= TableBorderStyle.Left;
                }
                if (tblrng._toCol == cell._fromCol)
                {
                    tableStyle.borderStyleType |= TableBorderStyle.Right;
                }
                //check horizontal
                //check vertical
            }
        }

        public void GetFillStyles( ExcelRangeBase cell, PdfCellStyleOverride cellStyle)
        {
            if (cell.Style.Fill.IsEmpty())
            {
                //Conditional Formating

                //Table
                var tables = cell.Worksheet.Tables.GetIntersectingRanges(cell);
                if (tables.Count > 0)
                {
                    var table = tables[0].Value;
                    var range = table.Range;

                    int tableRow = 0;
                    int tableCol = 0;
                    ExcelTableNamedStyle tableStyle;
                    if (table.TableStyle == TableStyles.Custom)
                    {
                        tableStyle = cell.Worksheet.Workbook.Styles.TableStyles[table.StyleName].As.TableStyle;
                    }
                    else
                    {
                        var tmpNode = table.WorkSheet.Workbook.StylesXml.CreateElement("c:tableStyle");
                        tableStyle = new ExcelTableNamedStyle(cell.Worksheet.Workbook.Styles.NameSpaceManager, tmpNode, cell.Worksheet.Workbook.Styles);
                        tableStyle.SetFromTemplate((TableStyles)table.TableStyle);
                    }
                    tableRow = cell._fromRow - range._fromRow;
                    tableCol = cell._fromCol - range._fromCol;
                    if (table.ShowHeader && tableRow == 0)
                    {
                        cellStyle.dxfFill = tableStyle.HeaderRow.Style.Fill;
                    }
                    else if (table.ShowTotal && range._toRow == cell._fromRow)
                    {
                        cellStyle.dxfFill = tableStyle.TotalRow.Style.Fill;
                    }
                    else if (table.ShowFirstColumn && tableCol == 0)
                    {
                        cellStyle.dxfFill = tableStyle.FirstColumn.Style.Fill;
                    }
                    else if (table.ShowLastColumn && range._toCol == cell._fromCol)
                    {
                        cellStyle.dxfFill = tableStyle.LastColumn.Style.Fill;
                    }
                    else if (table.ShowRowStripes)
                    {
                        cellStyle.dxfFill = (tableRow & 1) == 0 ? tableStyle.SecondRowStripe.Style.Fill : tableStyle.FirstRowStripe.Style.Fill;
                    }
                    else if (table.ShowColumnStripes)
                    {
                        cellStyle.dxfFill = (tableCol & 1) == 0 ? tableStyle.SecondColumnStripe.Style.Fill : tableStyle.FirstColumnStripe.Style.Fill;
                    }
                    else
                    {
                        cellStyle.dxfFill = tableStyle.WholeTable.Style.Fill;
                    }
                }
            }
            cellStyle.xfFill = cell.Style.Fill;
        }

        //Create cell.
        private void HandleCell(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, double x, double y, double width, double height, PdfCellStyleOverride CellStyle)
        {
            //We add empty cells for gridline calculation later. We just marked them for deletion by addng * to their name.
            string deleteMark = !cell.IsEmpty() || cell.Worksheet.ExistsStyleInner(cell._fromRow, cell._toCol) ? "" : "*";
            var cl0 = new PdfCellLayout(dictionaries, cell, CellStyle, x, y, width, height, 1, 1, 0, this);
            cl0.Name = cell.Address + deleteMark;
            cl0.Z = 1;
            AddCellContent(pageSettings, dictionaries, cell, x, y-height, width, height, 2);
            var border = HandleDiagonalBorders(cell, CellStyle, x, y, width, height);
            if (border != null)
            {
                border.InitDiagonalBorders(cell, width, height);
            }
        }

        //Create merged cell.
        private void HandleMergedCell(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, List<string> checkedMergedCells, double x, double y)
        {
            string mergeAddress = worksheet.MergedCells[cell.Start.Row, cell.Start.Column];
            if (!checkedMergedCells.Contains(mergeAddress))
            {
                double height = 0, width = 0;
                ExcelAddressBase address = new ExcelAddressBase(mergeAddress);
                for (int k = address._fromRow; k <= address._toRow; k++)
                {
                    height += UnitConversion.ExcelRowHeightToPoints(worksheet.Row(k).Height);
                }
                for (int l = address._fromCol; l <= address._toCol; l++)
                {
                    width += UnitConversion.ExcelColumnWidthToPoints(worksheet.Column(l).Width, ZeroCharWidth);
                }
                var mergedCell = AddChild(new PdfMergedCellLayout(dictionaries, worksheet.Cells[address._fromRow, address._fromCol], x, y, width, height));
                mergedCell.Name = cell.Address + "_m";
                mergedCell.Z = 5;
                AddCellContent(pageSettings, dictionaries, cell, x, y-height, width, height, 6);
                checkedMergedCells.Add(mergeAddress);
                var border = HandleDiagonalBorders(cell, null, x, y, width, height);
                if (border != null)
                {
                    border.InitDiagonalBorders(cell, width, height);
                    border.range = address.Address;
                }
            }
        }

        //Create content.
        private void AddCellContent(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, double x, double y, double width, double height, int zOrder)
        {
            if (!string.IsNullOrEmpty(cell.Text))
            {
                var cellContent = new PdfCellContentLayout(cell, pageSettings, x, y, width, height, 1, 1, 0, this, dictionaries);
                cellContent.Name = cell.Address + "_c";
                cellContent.Z = zOrder;
            }
        }

        //Create Edge borders.
        private PdfCellBorderLayout HandleEdgeBorders(ExcelRangeBase cell, PdfCellStyleOverride tableStyle, double x, double y, double width, double height)
        {
            bool edges = new[] { cell.Style.Border.Top.Style, cell.Style.Border.Bottom.Style, cell.Style.Border.Left.Style, cell.Style.Border.Right.Style }.All(s => s == ExcelBorderStyle.None);
            if (!edges)
            {
                var clb0 = new PdfCellBorderLayout(cell, null, x, y, width, height, 1, 1, 0, this);
                clb0.Name = cell.Address + "_b";
                clb0.Z = 7;
                return clb0;
            }
            if (tableStyle != null)
            {
                var clb0 = new PdfCellBorderLayout(cell, tableStyle, x, y, width, height, 1, 1, 0, this);
                clb0.Name = cell.Address + "_b";
                clb0.Z = 7;
                return clb0;
            }
            return null;
        }

        //Create Diagonal borders.
        private PdfCellBorderLayout HandleDiagonalBorders(ExcelRangeBase cell, PdfCellStyleOverride TableStyle, double x, double y, double width, double height)
        {
            bool diagonals = new[] { cell.Style.Border.Diagonal.Style }.All(s => s == ExcelBorderStyle.None);
            if (!diagonals)
            {
                var clb0 = new PdfCellBorderLayout(cell, TableStyle, x, y, width, height, 1, 1, 0, this);
                clb0.Name = cell.Address + "_b";
                clb0.Z = 7;
                return clb0;
            }
            return null;
        }

        //Create drawings.
        private void HandleDrawings(ExcelWorksheet worksheet)
        {
            foreach (var drawing in worksheet.Drawings) //NOT IMPLEMENTED
            {
                var drawLayout = AddChild(new PdfDrawingLayout(drawing, drawing.Position.X, drawing.Position.Y, drawing._width, drawing._height));
                drawLayout.Z = 10;
                drawLayout.Name = "Drawing " + drawing.Name;
            }
        }

        private double GetThemeFont0Width(ExcelWorksheet ws)
        {
            FontMeasurerTrueType fontMeasurerTrueType = new FontMeasurerTrueType();
            MeasurementFont font = new MeasurementFont();
            var ns = ws.Workbook.Styles.GetNormalStyle();
            font.FontFamily = ns.Style.Font.Name;
            font.Size = ns.Style.Font.Size;
            font.Style = MeasurementFontStyles.Regular;
            var result = fontMeasurerTrueType.MeasureText("0", font);
            return result.Width;
        }
    }
}
