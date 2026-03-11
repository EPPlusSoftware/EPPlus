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

        internal ConnectionPointsMiddle connectionPoints;

        private PointLines(DrawingBase renderer) : base(renderer)
        {
        }

        internal PointLines(DrawingBase renderer, BoundingBox parent, ConnectionPointsMiddle connectionPoints) : this(renderer)
        {
            Bounds = new BoundingBox();

            Bounds.Parent = parent;
            Bounds.Left = parent.Left;
            Bounds.Top = parent.Top;
            Bounds.Width = parent.Width;
            Bounds.Height = parent.Height;

            //Add connection points to render
            List<string> ptColors = new List<string> { "red", "green", "blue", "yellow" };
            for (int i = 0; i < connectionPoints.Points.Count; i++)
            {
                var cPoint = connectionPoints.Points[i];
                var cPointLine = new SvgRenderLineItem(renderer, Bounds);
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
