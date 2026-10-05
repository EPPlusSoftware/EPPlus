using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.Svg;
using System.Collections.Generic;
using System.Drawing;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class PointLines : ChartDrawingObjectWithBackground
    {
        internal List<LineRenderItem> RenderLines = new List<LineRenderItem>();

        internal ConnectionPointsMiddle ConnectionPoints;

        private List<string> ptColors = new List<string> { "red", "green", "blue", "yellow" };

        private BoundingBox parentBounds;

        internal override Color? DefaultFillColor => Color.Black;

        internal override Color? DefaultBorderColor => Color.Black;

        private PointLines(ChartRenderer cr) : base(cr)
        {
            Rectangle = new RectRenderItem(cr.Bounds);
        }

        internal PointLines(ChartRenderer cr, BoundingBox parent, ConnectionPointsMiddle connectionPoints) : this(cr)
        {
            parentBounds = parent;

            Rectangle.Parent = parent;
            ConnectionPoints = connectionPoints;

            UpdateLines();
        }

        internal void UpdateLines()
        {
            RenderLines.Clear();

            for (int i = 0; i < ConnectionPoints.Points.Count; i++)
            {
                var cPoint = ConnectionPoints.Points[i];
                var cPointLine = new LineRenderItem(Rectangle);
                cPointLine.X1 = 0;
                cPointLine.Y1 = 0;
                cPointLine.X2 = cPoint.X;
                cPointLine.Y2 = cPoint.Y;

                cPointLine.BorderWidth = 1;
                cPointLine.BorderColor = ptColors[i];
                RenderLines.Add(cPointLine);
            }
        }

        public override void AppendRenderItems(List<Transform> renderItems)
        {
            GroupRenderItem gItem = new GroupRenderItem(Rectangle);
            renderItems.Add(gItem);
            foreach (var line in RenderLines)
            {
                gItem.ChildObjects.Add(line);
            }
        }

        internal override Color? GetDefaultFillColor()
        {
            return DefaultFillColor;
        }

        internal override Color? GetDefaultBorderColor()
        {
            return DefaultBorderColor;
        }
    }
}
