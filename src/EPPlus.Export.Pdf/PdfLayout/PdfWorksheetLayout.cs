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
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using EPPlus.Graphics.Units;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
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
                    PdfCellBorderLayout border = HandleBorders(cell, x, y, width, height);
                    if (cell.Merge)
                    {
                        HandleMergedCell(worksheet, pageSettings, dictionaries, cell, checkedMergedCells, border, x, y);
                    }
                    else
                    {
                        HandleCell(pageSettings, dictionaries, cell, x, y, width, height);
                        if (border != null) border.InitDiagonalBorders(cell, 0, 0);
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

        //Create cell.
        private void HandleCell(PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, double x, double y, double width, double height)
        {
            //We add empty cells for gridline calculation later. We just marked them for deletion by addng * to their name.
            string deleteMark = !cell.IsEmpty() || cell.Worksheet.ExistsStyleInner(cell._fromRow, cell._toCol) ? "" : "*";
            var cl0 = new PdfCellLayout(dictionaries, cell, x, y, width, height, 1, 1, 0, this);
            cl0.Name = cell.Address + deleteMark;
            cl0.Z = 1;
            AddCellContent(pageSettings, dictionaries, cell, x, y-height, width, height, 2);
        }

        //Create merged cell.
        private void HandleMergedCell(ExcelWorksheet worksheet, PdfPageSettings pageSettings, PdfDictionaries dictionaries, ExcelRangeBase cell, List<string> checkedMergedCells, PdfCellBorderLayout border, double x, double y)
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
                if (border != null)
                {
                    border.InitDiagonalBorders(cell, width, height);
                    //This makes borders in merged cells draw double borders on the Left side and Top side of the merged cell. This makes some borders look the way they are not supposed to look. However this does makes diagonal borders cover the full merged cell. We want our cake and eat it so we need to redesign border handling.
                    border.Size = new Vector2(width, height);
                    border.LocalPosition = new Vector2(x, y - height);
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

        //Create borders.
        //TODO: Fix border by having having an array of border objects for each edge, but keep only one of the diagonal
        //loop all cells in a merged cell and set border style for each cell
        //calculate start and en position of diagonal
        //store it and use them when creating pdf objects
        private PdfCellBorderLayout HandleBorders(ExcelRangeBase cell, double x, double y, double width, double height)
        {
            bool allNone = new[] { cell.Style.Border.Top.Style, cell.Style.Border.Bottom.Style, cell.Style.Border.Left.Style, cell.Style.Border.Right.Style, cell.Style.Border.Diagonal.Style }.All(s => s == ExcelBorderStyle.None);
            if (!allNone)
            {
                var clb0 = new PdfCellBorderLayout(cell, x, y, width, height, 1, 1, 0, this);
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
