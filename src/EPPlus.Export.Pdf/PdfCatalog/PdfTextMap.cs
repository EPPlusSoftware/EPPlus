using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Graphics.Units;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfTextMap
    {
        public PdfTextMap(PdfRange Range)
        {
            var worksheet = Range.Range.Worksheet;
            List<string> checkedMergedCells = new List<string>();
            //int addedColumns = Range.ExtendColumns ? AddColumnsForNonWrappedText(worksheet) : 0;
            for (int row = Range.Range._fromRow; row <= Range.Range._toRow; row++)
            {
                if (worksheet.Row(row).Hidden) continue;
                //var height = UnitConversion.ExcelRowHeightToPoints(worksheet.Row(row).Height);
                //x = 0d;
                for (int col = Range.Range._fromCol; col <= Range.Range._toCol /*+ addedColumns*/; col++)
                {
                    if (worksheet.Column(col).Hidden) continue;
                    //var width = UnitConversion.ExcelColumnWidthToPoints(worksheet.Column(col).Width, ZeroCharWidth);
                    var cell = worksheet.Cells[row, col];

                    //get text

                    //Get Comments



                    //var cellStyle = new PdfCellStyle();
                    //GetFillStyles(cell, cellStyle);
                    //GetBorderStyles(cell, cellStyle);
                    //GetFontStyles(cell, cellStyle);
                    //PdfCellBorderLayout border = HandleEdgeBorders(cell, cellStyle, cell.Address, x, y, width, height);
                    //if (cell.Merge)
                    //{
                    //    HandleMergedCell(worksheet, pageSettings, dictionaries, cell, cellStyle, checkedMergedCells, x, y);
                    //}
                    //else
                    //{
                    //    HandleCell(pageSettings, dictionaries, cell, x, y, width, height, cellStyle);
                    //}
                    //if (border != null) border.InitEdgeBorders(cell);
                    //x += width;
                    //totalWidth = System.Math.Max(x, totalWidth);
                    //if (pageSettings.CommentsAndNotes != CommentsAndNotes.None)
                    //{
                    //    if (cell.Comment != null && cell.ThreadedComment == null)
                    //    {
                    //        dictionaries.CommentsAndNotes.Add(cell.Address, new PdfCommentsAndNotes(cell.Comment));
                    //    }
                    //    if (cell.ThreadedComment != null)
                    //    {
                    //        dictionaries.CommentsAndNotes.Add(cell.Address, new PdfCommentsAndNotes(cell.ThreadedComment));
                    //        PdfCommentsAndNotes.HasThreadedComment = true;
                    //    }
                    //}
                }
                //y -= height;
            }
            //HandleDrawings(worksheet);
            //Size = new Vector2(totalWidth, Math.Abs(y));
        }
    }
}
