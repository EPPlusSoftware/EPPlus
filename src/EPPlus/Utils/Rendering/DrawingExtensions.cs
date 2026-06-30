using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.ExternalReferences;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Utils
{
    internal static class DrawingExtensions
    {
        internal static BoundingBox GetBoundingBox(this ExcelDrawing drawing)
        {
            return new BoundingBox()
            {
                Left = 0,
                Top = 0,
                Width = drawing.GetPixelWidth().PixelToPoint(),
                Height = drawing.GetPixelHeight().PixelToPoint()
            };
        }

        internal static List<object> LoadSeriesValues(ExcelChart chart, string serieAddressInput, double[] numLiterals, string[] strLiterals)
        {
            string serieAddress = serieAddressInput;

            //Some addresses are split and within parenthesis
            if (serieAddressInput.StartsWith("("))
            {
                serieAddress = serieAddressInput.Trim('(', ')');
            }

            List<object> values = new List<object>();
            if (numLiterals != null)
            {
                values.AddRange(numLiterals.Select(x => (object)x));
            }
            else if (strLiterals != null)
            {
                values.AddRange(strLiterals.Select(x => (object)x));
            }
            else
            {
                if (string.IsNullOrEmpty(serieAddress))
                {
                    return null;
                }
                var address = new ExcelAddressBase(serieAddress);

                if (address.Addresses != null && address.Addresses.Count > 1)
                {
                    foreach (var splitAddress in address.Addresses)
                    {
                        FillValuesFromAddress(chart, splitAddress, ref values);
                    }
                }
                else
                {
                    FillValuesFromAddress(chart, address, ref values);
                }
            }
            return values;
        }

        internal static void FillValuesFromAddress(ExcelChart Chart, ExcelAddressBase address, ref List<object> values)
        {
            if (address.IsExternal)
            {
                var wb = Chart.WorkSheet.Workbook;
                var extWb = wb.ExternalLinks[address.ExternalReferenceIndex - 1] as ExcelExternalWorkbook;
                if (extWb != null)
                {
                    var wsName = address.WorkSheetName;
                    if (extWb.Package == null)
                    {
                        var extWs = extWb.CachedWorksheets[wsName];
                        FillExternalValues(extWs, address, ref values);
                    }
                    else
                    {
                        var ws = extWb.Package.Workbook.Worksheets[wsName];
                        FillInternalValues(ws, address, ref values);
                    }
                }
            }
            else
            {
                var wsName = address.WorkSheetName;

                if (string.IsNullOrEmpty(wsName))
                {
                    wsName = Chart.WorkSheet.Name;
                }

                var ws = Chart.WorkSheet.Workbook.Worksheets[wsName];
                FillInternalValues(ws, address, ref values);
            }
        }

        internal static void FillExternalValues(ExcelExternalWorksheet extWs, ExcelAddressBase address, ref List<object> values)
        {
            if (extWs != null)
            {
                for (int r = address.Start.Row; r <= address.End.Row; r++)
                {
                    for (int c = address.Start.Column; c <= address.End.Column; c++)
                    {
                        values.Add(extWs.CellValues[r, c].Value);
                    }
                }
            }
        }

        internal static void FillInternalValues(ExcelWorksheet ws, ExcelAddressBase address, ref List<object> values)
        {

            if (ws != null)
            {
                for (int r = address.Start.Row; r <= address.End.Row; r++)
                {
                    for (int c = address.Start.Column; c <= address.End.Column; c++)
                    {
                        values.Add(ws.Cells[r, c].Value);
                    }
                }
            }
        }


        internal static OffsetRectangle AsOffsetRectangle(this ExcelDrawingRectangle item)
        {
            return new OffsetRectangle
            {
                TopOffset = item.TopOffset,
                BottomOffset = item.BottomOffset,
                LeftOffset = item.LeftOffset,
                RightOffset = item.RightOffset
            };
        }
        internal static FillTile AsFillTile(this ExcelDrawingBlipFillTile fillTile)
        {
            return new FillTile
            {
                Alignment = (RectangleAlignment?)fillTile.Alignment,
                FlipMode = (TileFlipMode)fillTile.FlipMode,
                HorizontalOffset = fillTile.HorizontalOffset,
                VerticalOffset = fillTile.VerticalOffset,
                HorizontalRatio = fillTile.HorizontalRatio,
                VerticalRatio = fillTile.VerticalRatio
            };
        }
    }
}
