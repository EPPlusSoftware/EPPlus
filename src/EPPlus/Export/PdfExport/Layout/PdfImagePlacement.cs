using EPPlus.Export.Pdf.DocumentObjects;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Graphics.Units;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Export.PdfExport.TextMapping;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfImagePlacement
    {
        internal static void AddCellPicture(PdfPageLayout pageLayout, PdfCell map, double x, double y, double effectiveWidth, double rowHeight)
        {
            if (map.CellPictureBytes != null && map.CellPicturePixelWidth > 0 && map.CellPicturePixelHeight > 0)
            {
                const double inset = 1d;
                double availW = effectiveWidth - 2 * inset;
                double availH = rowHeight - 2 * inset;
                if (availW > 0 && availH > 0)
                {
                    double scale = System.Math.Min(availW / map.CellPicturePixelWidth, availH / map.CellPicturePixelHeight);
                    double dispW = map.CellPicturePixelWidth * scale;
                    double dispH = map.CellPicturePixelHeight * scale;
                    double imgX = x + inset + CellImageOffsetX(map.ContentAligmnet, availW - dispW);
                    double imgTop = y - inset - CellImageOffsetY(map.ContentAligmnet, availH - dispH);
                    pageLayout.AddChild(new PdfImageLayout(imgX, imgTop, dispW, dispH)
                    {
                        ImageBytes = map.CellPictureBytes,
                        Name = "CellImage_" + map.Name,
                    });
                }
            }
        }

        private static double CellImageOffsetX(PdfCellAlignmentData align, double slack)
        {
            if (slack <= 0d || align == null) return slack > 0d ? slack / 2d : 0d;
            switch (align.HorizontalAlignment)
            {
                case EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.General:
                case EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.Left:
                case EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.Fill:
                case EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.Justify:
                case EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.Distributed:
                    return 0d;
                case EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.Right:
                    return slack;
                default: // Center / CenterContinuous
                    return slack / 2d;
            }
        }

        private static double CellImageOffsetY(PdfCellAlignmentData align, double slack)
        {
            if (slack <= 0d || align == null) return slack > 0d ? slack / 2d : 0d;
            switch (align.VerticalAlignment)
            {
                case EPPlus.Export.Pdf.Enums.ExcelVerticalAlignment.Top:
                    return 0d;
                case EPPlus.Export.Pdf.Enums.ExcelVerticalAlignment.Bottom:
                    return slack;
                default: // Center / Distributed / Justify
                    return slack / 2d;
            }
        }

        internal static Pages PrecomputeImages(PdfPageSettings pageSettings, PdfRange range, Pages pdfPages, List<PdfDrawing> drawings, double zeroCharWidth)
        {
            if (drawings == null || drawings.Count == 0) return pdfPages;
            var colPrefix = new double[range.ColWidths.Count + 1];
            for (int i = 0; i < range.ColWidths.Count; i++)
                colPrefix[i + 1] = colPrefix[i] + range.ColWidths[i];
            var rowPrefix = new double[range.RowHeights.Count + 1];
            for (int i = 0; i < range.RowHeights.Count; i++)
                rowPrefix[i + 1] = rowPrefix[i] + range.RowHeights[i].Height;
            var ws = range.Range.Worksheet;
            double rangeOriginX = 0d;
            for (int c = 1; c < range.Range._fromCol; c++)
                rangeOriginX += ws.Column(c).Hidden ? 0d : UnitConversion.ExcelColumnWidthToPoints(ws.Column(c).Width, zeroCharWidth);
            double rangeOriginY = 0d;
            for (int r = 1; r < range.Range._fromRow; r++)
                rangeOriginY += ws.Row(r).Hidden ? 0d : ws.Row(r).Height;
            for (int i = 0; i < pdfPages.Page.Length; i++)
                pdfPages.Page[i] = PrecomputePageImages(pageSettings, range, pdfPages.Page[i], drawings, colPrefix, rowPrefix, rangeOriginX, rangeOriginY);
            return pdfPages;
        }
        private static double PointsFromEmu(long emu) => emu / (double)ExcelDrawing.EMU_PER_POINT;
        private static double PointsFromPixels(double pixels) => pixels * ExcelDrawing.EMU_PER_PIXEL / (double)ExcelDrawing.EMU_PER_POINT;
        private static double ColumnEdge(double[] colPrefix, int localCol) => colPrefix[Math.Max(0, Math.Min(localCol, colPrefix.Length - 1))];
        private static double RowEdge(double[] rowPrefix, int localRow) => rowPrefix[Math.Max(0, Math.Min(localRow, rowPrefix.Length - 1))];
        private static Page PrecomputePageImages(PdfPageSettings pageSettings, PdfRange range, Page page, List<PdfDrawing> drawings, double[] colPrefix, double[] rowPrefix, double rangeOriginX, double rangeOriginY)
        {
            page.Images = new List<ImageDrawInfo>();
            int fromCol = range.Range._fromCol;
            int fromRow = range.Range._fromRow;

            double pageAbsLeft = colPrefix[page.FromColumn - fromCol];
            double pageAbsRight = colPrefix[page.ToColumn - fromCol + 1];
            double pageAbsTop = rowPrefix[page.FromRow - fromRow];
            double pageAbsBottom = rowPrefix[page.ToRow - fromRow + 1];

            double contentStartX = GetOriginX(pageSettings, page) + page.HeadingWidth + page.PrintTitleWidth;
            double contentStartY = GetOriginY(pageSettings, page) - page.HeadingHeight - page.PrintTitleHeight;

            foreach (var drawing in drawings)
            {
                if (!PdfImageXObject.CanEmbed(drawing.ImageBytes)) continue;
                var pic = drawing.Picture;
                double imgLeft, imgTop, imgRight, imgBottom;
                if (pic.From != null)
                {
                    int imgColLocal = (pic.From.Column + 1) - fromCol;
                    int imgRowLocal = (pic.From.Row + 1) - fromRow;
                    if (imgColLocal < 0 || imgColLocal >= colPrefix.Length - 1) continue;
                    if (imgRowLocal < 0 || imgRowLocal >= rowPrefix.Length - 1) continue;
                    imgLeft = colPrefix[imgColLocal] + PointsFromEmu(pic.From.ColumnOff);
                    imgTop = rowPrefix[imgRowLocal] + PointsFromEmu(pic.From.RowOff);
                    if (pic.To != null)
                    {
                        imgRight = ColumnEdge(colPrefix, (pic.To.Column + 1) - fromCol) + PointsFromEmu(pic.To.ColumnOff);
                        imgBottom = RowEdge(rowPrefix, (pic.To.Row + 1) - fromRow) + PointsFromEmu(pic.To.RowOff);
                    }
                    else
                    {
                        imgRight = imgLeft + PointsFromPixels(pic.GetPixelWidth());
                        imgBottom = imgTop + PointsFromPixels(pic.GetPixelHeight());
                    }
                }
                else if (pic.Position != null && pic.Size != null)
                {
                    imgLeft = PointsFromEmu(pic.Position.X) - rangeOriginX;
                    imgTop = PointsFromEmu(pic.Position.Y) - rangeOriginY;
                    imgRight = imgLeft + PointsFromEmu(pic.Size.Width);
                    imgBottom = imgTop + PointsFromEmu(pic.Size.Height);
                }
                else
                {
                    continue;
                }
                double width = imgRight - imgLeft;
                double height = imgBottom - imgTop;

                if (imgLeft >= pageAbsRight || imgRight <= pageAbsLeft) continue;
                if (imgTop >= pageAbsBottom || imgBottom <= pageAbsTop) continue;

                double x = contentStartX + (imgLeft - pageAbsLeft);
                double y = contentStartY - (imgTop - pageAbsTop);

                string hlUri = null, hlSheet = null, hlToolTip = null;
                int hlRow = 0, hlCol = 0;
                var picLink = pic.Hyperlink;
                if (picLink != null)
                {
                    string tip = (picLink as ExcelHyperLink)?.ToolTip;
                    if (picLink is ExcelHyperLink peh && !string.IsNullOrEmpty(peh.ReferenceAddress))
                    {
                        if (PdfTextMap.TryResolveInternalLink(range.Range.Worksheet, peh.ReferenceAddress, out var ts, out var tr, out var tc))
                        {
                            hlSheet = ts; hlRow = tr; hlCol = tc;
                            hlToolTip = !string.IsNullOrEmpty(tip) ? tip : peh.ReferenceAddress;
                        }
                    }
                    else
                    {
                        hlUri = PdfTextMap.NormalizeExternalTarget(picLink.OriginalString);
                        if (hlUri != null)
                            hlToolTip = !string.IsNullOrEmpty(tip) ? tip : picLink.OriginalString;
                    }
                }

                page.Images.Add(new ImageDrawInfo
                {
                    X = x,
                    Y = y,
                    Width = width,
                    Height = height,
                    ImageBytes = drawing.ImageBytes,
                    Name = "Image_" + pic.Name,
                    HyperlinkUri = hlUri,
                    HyperlinkTargetSheet = hlSheet,
                    HyperlinkTargetRow = hlRow,
                    HyperlinkTargetColumn = hlCol,
                    HyperlinkToolTip = hlToolTip,
                });
            }
            return page;
        }
    }
}
