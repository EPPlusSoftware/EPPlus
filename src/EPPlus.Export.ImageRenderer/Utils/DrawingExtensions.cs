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
                Left = drawing.GetPixelLeft(),
                Top = drawing.GetPixelTop(),
                Width = drawing.GetPixelWidth(),
                Height = drawing.GetPixelHeight()
            };
        }
    }
}
