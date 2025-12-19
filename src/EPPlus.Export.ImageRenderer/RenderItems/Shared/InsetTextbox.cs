using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    /// <summary>
    /// Calculated textbox position from a shape
    /// That contains the textbody
    /// </summary>
    internal class InsetTextbox: BoundingBox
    {
        TextBody textBody;

        //internal InsetTextbox(double l, double t, double width, double height)
        //{
        //    textBody = new TextBody(this);
        //}
    }
}
