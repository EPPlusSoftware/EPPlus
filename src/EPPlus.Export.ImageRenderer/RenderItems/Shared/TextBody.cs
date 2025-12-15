using EPPlus.Export.ImageRenderer.Text;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlus.Export.ImageRenderer.RenderItems.Shared
{
    internal class TextBody : FontWrapContainer
    {
        List<FontWrapContainer> Runs = new List<FontWrapContainer>();

        public TextBody(FontMeasurerTrueType txtMeasurer, Rect parent, bool initDefaults = true) : base(txtMeasurer, parent, initDefaults)
        {

        }

        //public TextBody(FontMeasurerTrueType txtMeasurer, Rect parent, bool initDefaults = true) : base(txtMeasurer, parent, initDefaults)
        //{

        //}
    }
}
