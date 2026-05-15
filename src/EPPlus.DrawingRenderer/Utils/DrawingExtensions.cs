using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing;

namespace EPPlus.Export.ImageRenderer.Utils
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
    }
}
