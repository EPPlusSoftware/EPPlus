using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class PointLines : DrawingObject
    {
        internal List<SvgRenderLineItem> RenderLines = new List<SvgRenderLineItem>();

        internal ConnectionPointsMiddle ConnectionPoints;

        private List<string> ptColors = new List<string> { "red", "green", "blue", "yellow" };

        private BoundingBox parentBounds;

        private PointLines(DrawingBase renderer) : base(renderer)
        {
        }

        internal PointLines(DrawingBase renderer, BoundingBox parent, ConnectionPointsMiddle connectionPoints) : this(renderer)
        {
            parentBounds = parent;

            Bounds = new BoundingBox();

            Bounds.Parent = parent;
            //Bounds.Left = parent.Left;
            //Bounds.Top = parent.Top;
            //Bounds.Width = parent.Width;
            //Bounds.Height = parent.Height;
            ConnectionPoints = connectionPoints;

            UpdateLines();
        }

        internal void UpdateLines()
        {
            RenderLines.Clear();

            for (int i = 0; i < ConnectionPoints.Points.Count; i++)
            {
                var cPoint = ConnectionPoints.Points[i];
                var cPointLine = new SvgRenderLineItem(DrawingRenderer, Bounds);
                cPointLine.X1 = 0;
                cPointLine.Y1 = 0;
                cPointLine.X2 = cPoint.X;
                cPointLine.Y2 = cPoint.Y;

                cPointLine.BorderWidth = 1;
                cPointLine.BorderColor = ptColors[i];
                RenderLines.Add(cPointLine);
            }
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            SvgGroupItem gItem = new SvgGroupItem(DrawingRenderer, Bounds);
            renderItems.Add(gItem);
            foreach (var line in RenderLines)
            {
                renderItems.Add(line);
            }
            SvgEndGroupItem endGItem = new SvgEndGroupItem(DrawingRenderer, Bounds);
            renderItems.Add(endGItem);
        }
    }
}
