using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlus.Export.ImageRenderer.Svg
{
    internal class DrawingItemForTesting : DrawingBase
    {
        /// <summary>
        /// List of items to be added to this basic container
        /// </summary>
        internal List<DrawingObject> ExternalRenderItems = new List<DrawingObject>();
        internal List<DrawingObjectNoBounds> ExternalRenderItemsNoBounds = new List<DrawingObjectNoBounds>();

        internal DrawingItemForTesting(BoundingBox bounds) : base()
        {
            Bounds = bounds;
            SvgRenderRectItem bg = new SvgRenderRectItem(this, Bounds);

            bg.Width = Bounds.Width;
            bg.Height = Bounds.Height;
            bg.FillColor = "blue";
            bg.FillOpacity = 0.2d;

            RenderItems.Add(bg);
        }

        public void Render(StringBuilder sb)
        {
            RenderItems.Add(new SvgGroupItem(this));
            foreach(var item in ExternalRenderItemsNoBounds)
            {
                item.AppendRenderItems(RenderItems);
            }
            foreach(var item in ExternalRenderItems)
            {
                item.AppendRenderItems(RenderItems);
            }
            RenderItems.Add(new SvgEndGroupItem(this, Bounds));

            sb.Append($"<svg width=\"{Bounds.Width.PointToPixelString()}\" height=\"{Bounds.Height.PointToPixelString()}\" xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" xml:space=\"default\" Overflow=\"Hidden\">");

            foreach (var item in RenderItems)
            {
                item.Render(sb);
            }

            sb.Append("</svg>");
        }
    }
}
