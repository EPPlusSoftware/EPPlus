using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Independent.Shared
{
    internal abstract class IndependentTextBox : DrawingObject
    {
        internal TextBodyItem TextBody;
        /// <summary>
        /// A textbox that is not reliant on epplus or drawings to function
        /// </summary>
        /// <param name="bounds">The bounds of this textbox</param>
        public IndependentTextBox(DrawingBase baseObj, BoundingBox bounds) : base(baseObj)
        {
            Bounds = bounds;
            TextBody = CreateTextBody(baseObj, Bounds);
        }

        /// <summary>
        /// Each file format defines its own paragraph
        /// </summary>
        /// <returns></returns>
        internal abstract TextBodyItem CreateTextBody(DrawingBase obj, BoundingBox parent);
    }
}
