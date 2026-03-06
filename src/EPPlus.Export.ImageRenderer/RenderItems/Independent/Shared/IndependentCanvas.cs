using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Independent.Shared
{
    internal abstract class IndependentCanvas : RenderItemIndependent
    {
        internal protected IndependentRect Background;

        protected IndependentCanvas(BoundingBox canvasBounds, Color bgColor)
        {
            Bounds = canvasBounds;
            Background = CreateBackground(canvasBounds);
            Background.SetFillColor(bgColor);
        }

        internal void SetBgColor(Color bgColor)
        {
            Background.SetFillColor(bgColor);
        }

        internal abstract IndependentRect CreateBackground(BoundingBox backgroundBounds);
    }
}
