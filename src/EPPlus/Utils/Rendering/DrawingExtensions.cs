using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Fill;
using OfficeOpenXml.Drawing.Theme;

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
                Alignment = (RectangleAlignment)fillTile.Alignment,
                FlipMode = (TileFlipMode)fillTile.FlipMode,
                HorizontalOffset = fillTile.HorizontalOffset,
                VerticalOffset = fillTile.VerticalOffset,
                HorizontalRatio = fillTile.HorizontalRatio,
                VerticalRatio = fillTile.VerticalRatio
            };
        }
    }
}
