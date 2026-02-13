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
                Width = drawing.GetPixelWidth(),
                Height = drawing.GetPixelHeight()
            };
        }
    }
}
