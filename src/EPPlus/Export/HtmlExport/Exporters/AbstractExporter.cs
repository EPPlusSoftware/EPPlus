using OfficeOpenXml.Drawing;
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
    }
}
