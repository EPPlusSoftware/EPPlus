using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.Svg;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class PointLines : ChartDrawingObject
    {
        internal List<LineRenderItem> RenderLines = new List<LineRenderItem>();

        internal ConnectionPointsMiddle ConnectionPoints;

        private List<string> ptColors = new List<string> { "red", "green", "blue", "yellow" };

        private BoundingBox parentBounds;

        private PointLines(ChartRenderer cr) : base(cr)
        {
        }

        internal PointLines(ChartRenderer cr, BoundingBox parent, ConnectionPointsMiddle connectionPoints) : this(cr)
        {
            parentBounds = parent;

            Rectangle.Bounds.Parent = parent;
            ConnectionPoints = connectionPoints;

            UpdateLines();
        }

        internal void UpdateLines()
        {
            RenderLines.Clear();

            for (int i = 0; i < ConnectionPoints.Points.Count; i++)
            {
                var cPoint = ConnectionPoints.Points[i];
                var cPointLine = new LineRenderItem(Rectangle.Bounds);
                cPointLine.X1 = 0;
                cPointLine.Y1 = 0;
                cPointLine.X2 = cPoint.X;
                cPointLine.Y2 = cPoint.Y;

                cPointLine.BorderWidth = 1;
                cPointLine.BorderColor = ptColors[i];
                RenderLines.Add(cPointLine);
            }
        }

        public override void AppendRenderItems(List<RenderItem> renderItems)
        {
            GroupRenderItem gItem = new GroupRenderItem(Rectangle.Bounds);
            renderItems.Add(gItem);
            foreach (var line in RenderLines)
            {
                gItem.RenderItems.Add(line);
            }
        }
    }
}
