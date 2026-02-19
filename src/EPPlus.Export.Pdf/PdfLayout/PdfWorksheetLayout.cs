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
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
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
            int addedColumns = AddColumnsForNonWrappedText(worksheet);
            for(int row = 1; row<= worksheet.Dimension._toRow; row++)
            {
                if(worksheet.Row(row).Hidden) continue;
                var height = UnitConversion.ExcelRowHeightToPoints(worksheet.Row(row).Height);
                x = 0d;
                for (int col = 1; col <= worksheet.Dimension._toCol + addedColumns; col++)
                {
                    if(worksheet.Column(col).Hidden) continue;
                    var width = UnitConversion.ExcelColumnWidthToPoints(worksheet.Column(col).Width, ZeroCharWidth);
                    var cell = worksheet.Cells[row, col];
                    var cellStyle = new PdfCellStyle();
                    GetFillStyles(cell, cellStyle);
                    GetBorderStyles(cell, cellStyle);
                    GetFontStyles(cell, cellStyle);
                    PdfCellBorderLayout border = HandleEdgeBorders(cell, cellStyle, x, y, width, height);
                    if (cell.Merge)
                    {
                        HandleMergedCell(worksheet, pageSettings, dictionaries, cell, cellStyle, checkedMergedCells, x, y);
                    }
                    else
                    {
                        HandleCell(pageSettings, dictionaries, cell, x, y, width, height, cellStyle);
                    }
                    if (border != null) border.InitEdgeBorders(cell);
                    x += width;
                    totalWidth = System.Math.Max( x, totalWidth);
                }
                y -= height;
            }
            HandleDrawings(worksheet);
            Size = new Vector2(totalWidth, y);
        }

        /// <summary>
        /// Check if we need to add additional columns to accomodate text that is not wrapped and overlaps other cell.s
        /// </summary>
        /// <param name="ws">The worksheet to check.</param>
        /// <returns>The number of columns to add.</returns>
        private int AddColumnsForNonWrappedText(ExcelWorksheet ws)
        {
            double columnWidth = UnitConversion.ExcelColumnWidthToPoints(ws.Column(ws.Dimension._toCol).Width, ZeroCharWidth);
            int columnsToAdd = 0;
            FontMeasurerTrueType fontMeasurerTrueType = new FontMeasurerTrueType();
            MeasurementFont font = new MeasurementFont();
            double textLength = 0;
            for (int row = 1; row <= ws.Dimension._toRow; row++)
            {
                var cell = ws.Cells[row, ws.Dimension._toCol];
                if ((!string.IsNullOrEmpty(cell.Text) || cell.RichText.Count > 0) && !cell.Style.WrapText && !cell.Merge)
                {
                    if (cell.IsRichText)
                    {
                        foreach (var rt in cell.RichText)
                        {
                            font.FontFamily = rt.FontName;
                            font.Size = (float)rt.Size;
                            font.Style = ((cell.Style.Font.Bold ? MeasurementFontStyles.Bold : 0) |
                                          (cell.Style.Font.Italic ? MeasurementFontStyles.Italic : 0) |
                                          (cell.Style.Font.Strike ? MeasurementFontStyles.Strikeout : 0) |
                                          (cell.Style.Font.UnderLine ? MeasurementFontStyles.Underline : 0))
                                          switch
                            {
                                0 => MeasurementFontStyles.Regular,
                                var s => s
                            };
                            var result = fontMeasurerTrueType.MeasureText(rt.Text, font);
                            textLength += result.Width;
                        }
                    }
                    else
                    {
                        font.FontFamily = cell.Style.Font.Name;
                        font.Size = (float)cell.Style.Font.Size;
                        font.Style = ((cell.Style.Font.Bold ? MeasurementFontStyles.Bold : 0) |
                                      (cell.Style.Font.Italic ? MeasurementFontStyles.Italic : 0) |
                                      (cell.Style.Font.Strike ? MeasurementFontStyles.Strikeout : 0) |
                                      (cell.Style.Font.UnderLine ? MeasurementFontStyles.Underline : 0))
                                      switch
                        {
                            0 => MeasurementFontStyles.Regular,
                            var s => s
                        };
                        var result = fontMeasurerTrueType.MeasureText(cell.Text, font);
                        textLength += result.Width;
                    }
                    //loop next col width until text is included
                    //increase columns to add
                    while (textLength > columnWidth)
                    {
                        columnsToAdd++;
                        columnWidth += UnitConversion.ExcelColumnWidthToPoints(ws.Column(ws.Dimension._toCol + columnsToAdd).Width, ZeroCharWidth);
                    }
                }
            }
            return columnsToAdd;
        }

        //Get Border styles from cell, tables and conditional formatting. TODO: Add support for Conditional formatting.
        private void GetBorderStyles(ExcelRangeBase cell, PdfCellStyle cellStyle)
        {

            /* Kika på varje del av border top bottom left right
             * om cell har top använd den
             * om cell inte har border, gå igenom prio ordning på tabell borders och använd WholeTable om null.
             * om cellein i tabellen är i mitten av tabellen använd horizontal och vertical border istället.
             * I fallet top så ska om vi är i header row eller om cell fromrow är samma som table fromrow så ska top border vara top border. annar är det horizontal som gäller
             * Glöm ej vertical border om fromcol är samma som tabell fromcol.
             */
            if (cell != null)
            {
                cellStyle.xfTop = cell.Style.Border.Top;
                cellStyle.xfBottom = cell.Style.Border.Bottom;
                cellStyle.xfLeft = cell.Style.Border.Left;
                cellStyle.xfRight = cell.Style.Border.Right;
                var tables = cell.Worksheet.Tables.GetIntersectingRanges(cell);
                if (tables.Count > 0)
                {
                    var table = tables[0].Value;
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
                    cellStyle.dxfTop = PdfCellBorderLayout.GetTopBorderItem(cell, cellStyle.xfTop, table, tableStyle);
                    cellStyle.dxfBottom = PdfCellBorderLayout.GetBottomBorderItem(cell, cellStyle.xfBottom, table, tableStyle);
                    cellStyle.dxfLeft = PdfCellBorderLayout.GetLeftBorderItem(cell, cellStyle.xfLeft, table, tableStyle);
                    cellStyle.dxfRight = PdfCellBorderLayout.GetRightBorderItem(cell, cellStyle.xfRight, table, tableStyle);
                }
            }
        }

        //Get Fill style from cell, tables and conditional formatting. TODO: Add support for Conditional formatting.
        public void GetFillStyles( ExcelRangeBase cell, PdfCellStyle cellStyle)
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
                        cellStyle.dxfFill = (tableCol & 1) != 0 ? tableStyle.SecondColumnStripe.Style.Fill : tableStyle.FirstColumnStripe.Style.Fill;
                    }
                    else
                    {
                        cellStyle.dxfFill = tableStyle.WholeTable.Style.Fill;
                    }
                }
            }
            cellStyle.xfFill = cell.Style.Fill;
        }

        //Get Font style from cell, tables and conditional formatting. TODO: Add support for Conditional formatting.
        public void GetFontStyles(ExcelRangeBase cell, PdfCellStyle cellStyle)
        {
            var tables = cell.Worksheet.Tables.GetIntersectingRanges(cell);
            if (tables.Count > 0)
            {
                var table = tables[0].Value;
                var range = table.Range;
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
                int tableRow = cell._fromRow - range._fromRow;
                int tableCol = cell._fromCol - range._fromCol;
                var font = tableStyle.WholeTable.Style.Font;
                if (table.ShowHeader && tableRow == 0)
                {
                    if (tableStyle.HeaderRow.Style.Font.HasValue)
                    {
                        font = tableStyle.HeaderRow.Style.Font;
                    }
                    if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Font.HasValue)
                    {
                        font = tableStyle.FirstHeaderCell.Style.Font;
                    }
                    if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Font.HasValue)
                    {
                        font = tableStyle.LastHeaderCell.Style.Font;
                    }
                }
                else if (table.ShowTotal && cell._fromRow == range._toRow)
                {
                    if (tableStyle.TotalRow.Style.Font.HasValue)
                    {
                        font = tableStyle.TotalRow.Style.Font;
                    }
                    if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Font.HasValue)
                    {
                        font = tableStyle.FirstTotalCell.Style.Font;
                    }
                    if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Font.HasValue)
                    {
                        font = tableStyle.LastTotalCell.Style.Font;
                    }
                }
                else
                {
                    if (table.ShowColumnStripes && (tableCol & 1) == 0)
                    {
                        font = tableStyle.FirstColumnStripe.Style.Font;
                    }
                    if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Top.HasValue && (tableCol & 1) != 0)
                    {
                        font = tableStyle.SecondColumnStripe.Style.Font;
                    }
                    if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Font.HasValue && (tableRow & 1) != 0)
                    {
                        font = tableStyle.FirstRowStripe.Style.Font;
                    }
                    if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Font.HasValue && (tableRow & 1) == 0)
                    {
                        font = tableStyle.SecondRowStripe.Style.Font;
                    }
                    if (table.ShowLastColumn && tableStyle.LastColumn.Style.Font.HasValue && cell._fromCol == range._toCol)
                    {
                        font = tableStyle.LastColumn.Style.Font;
                    }
                    if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Font.HasValue && tableCol == range._toCol)
                    {
                        font = tableStyle.FirstColumn.Style.Font;
                    }
                }
                cellStyle.dxfFont = font;
            }
        }

        //Create cell.
        private void HandleCell(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, double x, double y, double width, double height, PdfCellStyle CellStyle)
        {
            //We add empty cells for gridline calculation later. We just marked them for deletion by addng * to their name.
            string deleteMark = !cell.IsEmpty() || cell.Worksheet.ExistsStyleInner(cell._fromRow, cell._toCol) ? "" : "*";
            var cl0 = new PdfCellLayout(dictionaries, cell, CellStyle, x, y, width, height, 1, 1, 0, this);
            cl0.Name = cell.Address + deleteMark;
            cl0.Z = 1;
            AddCellContent(pageSettings, dictionaries, cell, CellStyle, x, y-height, width, height, 2);
            var border = HandleDiagonalBorders(cell, CellStyle, x, y, width, height);
            if (border != null)
            {
                border.InitDiagonalBorders(cell, width, height);
            }
        }

        //Create merged cell.
        private void HandleMergedCell(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, PdfCellStyle CellStyle, List<string> checkedMergedCells, double x, double y)
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
                var mergedCell = AddChild(new PdfMergedCellLayout(dictionaries, worksheet.Cells[address._fromRow, address._fromCol], CellStyle, x, y, width, height));
                mergedCell.Name = cell.Address + "_m";
                mergedCell.Z = 5;
                AddCellContent(pageSettings, dictionaries, cell, CellStyle, x, y-height, width, height, 6);
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
        private void AddCellContent(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, PdfCellStyle CellStyle, double x, double y, double width, double height, int zOrder)
        {
            if (!string.IsNullOrEmpty(cell.Text))
            {
                var cellContent = new PdfCellContentLayout(cell, CellStyle, pageSettings, x, y, width, height, 1, 1, 0, this, dictionaries);
                cellContent.Name = cell.Address + "_c";
                cellContent.Z = zOrder;
            }
        }

        //Create Edge borders.
        private PdfCellBorderLayout HandleEdgeBorders(ExcelRangeBase cell, PdfCellStyle tableStyle, double x, double y, double width, double height)
        {
            bool edges = new[] { cell.Style.Border.Top.Style, cell.Style.Border.Bottom.Style, cell.Style.Border.Left.Style, cell.Style.Border.Right.Style }.All(s => s == ExcelBorderStyle.None);
            bool edges2 = new[] { tableStyle.dxfTop, tableStyle.dxfBottom, tableStyle.dxfLeft, tableStyle.dxfRight }.Any(s => s != null && s.HasValue);
            if (!edges || edges2)
            {
                var clb0 = new PdfCellBorderLayout(cell, tableStyle, x, y, width, height, 1, 1, 0, this);
                clb0.Name = cell.Address + "_b";
                clb0.Z = 7;
                return clb0;
            }
            return null;
        }

        //Create Diagonal borders.
        private PdfCellBorderLayout HandleDiagonalBorders(ExcelRangeBase cell, PdfCellStyle TableStyle, double x, double y, double width, double height)
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

        //Get the width of the character 0 from the current themes default font style. It's used to calculate width of cells.
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
