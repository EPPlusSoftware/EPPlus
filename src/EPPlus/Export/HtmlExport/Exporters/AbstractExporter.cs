using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class AbstractExporter
    {
        protected readonly string TableClass = "epplus-table";
        internal List<HtmlImage> _rangePictures = null;
        protected readonly CellDataWriter _cellDataWriter = new CellDataWriter();

        internal void LoadRangeImages(List<ExcelRangeBase> ranges)
        {
            if (_rangePictures != null)
            {
                return;
            }
            _rangePictures = new List<HtmlImage>();
            //Render in-cell images.
            foreach (var worksheet in ranges.Select(x => x.Worksheet).Distinct())
            {
                foreach (var d in worksheet.Drawings)
                {
                    if (d is ExcelPicture p)
                    {
                        p.GetFromBounds(out int fromRow, out int fromRowOff, out int fromCol, out int fromColOff);
                        p.GetToBounds(out int toRow, out int toRowOff, out int toCol, out int toColOff);

                        _rangePictures.Add(new HtmlImage()
                        {
                            WorksheetId = worksheet.PositionId,
                            Picture = p,
                            FromRow = fromRow,
                            FromRowOff = fromRowOff,
                            FromColumn = fromCol,
                            FromColumnOff = fromColOff,
                            ToRow = toRow,
                            ToRowOff = toRowOff,
                            ToColumn = toCol,
                            ToColumnOff = toColOff
                        });
                    }
                }
            }
        }

        protected string GetCellText(ExcelRangeBase cell, HtmlExportSettings settings)
        {
            if (cell.IsRichText)
            {
                return cell.RichText.HtmlText;
            }
            else
            {
                return ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, settings.Culture);
            }
        }

        protected string GetImageCellClassName(HtmlImage image, HtmlExportSettings settings)
        {
            return image == null && settings.Pictures.Position != ePicturePosition.Absolute ? "" : settings.StyleClassPrefix + "image-cell";
        }

        protected HtmlImage GetImage(int worksheetId, int row, int col)
        {
            if (_rangePictures == null) return null;
            foreach (var p in _rangePictures)
            {
                if (p.FromRow == row - 1 && p.FromColumn == col - 1 && p.WorksheetId == worksheetId)
                {
                    return p;
                }
            }
            return null;
        }
    }
}
