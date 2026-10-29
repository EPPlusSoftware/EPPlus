using EPPlus.Export.Pdf.Layout;
using OfficeOpenXml.Export.PdfExport.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfLinkLayout
    {
        internal static bool CanTightenLink(PdfCellContentLayout text) =>
                text?.TextLines != null && text.TextLines.Count > 0 &&
                text.CellAlignmentData != null &&
                !text.CellAlignmentData.IsVertical &&
                text.CellAlignmentData.TextRotation == 0;

        internal static LinkDrawInfo CreateLinkDrawInfo(double x, double topY, double width, double height, PdfCell map)
        {
            return new LinkDrawInfo
            {
                X = x,
                Y = topY - height,
                Width = width,
                Height = height,
                Uri = map.HyperlinkUri,
                TargetSheet = map.HyperlinkTargetSheet,
                TargetRow = map.HyperlinkTargetRow,
                TargetColumn = map.HyperlinkTargetCol,
                ToolTip = map.HyperlinkToolTip
            };
        }

        internal static void AddTextLinks(PdfPageLayout pageLayout, PdfCellContentLayout text, double clipLeft, double clipRight, string uri, string targetSheet, int targetRow, int targetCol, string toolTip, double padX = 1.0d, double padY = 1.5d)
        {
            double line0Width = text.TextLines[0].Width;
            double originX = text.LocalPosition.X;   // aligned left of line 0
            double baseY = text.LocalPosition.Y;   // baseline of line 0
            for (int k = 0; k < text.TextLines.Count; k++)
            {
                var line = text.TextLines[k];
                double offX = 0d;                    // mirrors PdfContentStream.AddText's per-line offset
                switch (text.CellAlignmentData.HorizontalAlignment)
                {
                    case EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.Right:
                        offX = line0Width - line.Width; break;
                    case EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.Center:
                    case EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.CenterContinuous:
                    case EPPlus.Export.Pdf.Enums.ExcelHorizontalAlignment.Distributed:
                        offX = (line0Width - line.Width) / 2d; break;
                }
                double lineLeft = originX + offX;
                double boxLeft = System.Math.Max(lineLeft, clipLeft) - padX;
                double boxRight = System.Math.Min(lineLeft + line.Width, clipRight) + padX;
                if (line.Width > 0d && boxRight > boxLeft)
                {
                    pageLayout.Links.Add(new LinkDrawInfo
                    {
                        X = boxLeft,
                        Y = baseY - line.LargestDescent - padY,
                        Width = boxRight - boxLeft,
                        Height = line.LargestAscent + line.LargestDescent + 2 * padY,
                        Uri = uri,
                        TargetSheet = targetSheet,
                        TargetRow = targetRow,
                        TargetColumn = targetCol,
                        ToolTip = toolTip
                    });
                }
                baseY -= (line.LargestAscent + line.LargestDescent);
            }
        }
    }
}
