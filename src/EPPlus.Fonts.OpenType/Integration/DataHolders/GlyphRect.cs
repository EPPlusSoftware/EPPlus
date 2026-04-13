using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    internal class GlyphRect
    {
        //BoundingRectangle BoundingRectFontDesign;
        ushort glyphIndex;
        internal double advanceWidth;

        //Debug var only?
        string fontName;

        internal GlyphRect(ushort glyphIndex, double advanceWidth, string fontName)
        {
            this.glyphIndex = glyphIndex;
            this.advanceWidth = advanceWidth;
            this.fontName = fontName;
        }
        //BoundingRectangle CalculateBoundingRect(double fontSize)
        //{

        //}
    }
}
