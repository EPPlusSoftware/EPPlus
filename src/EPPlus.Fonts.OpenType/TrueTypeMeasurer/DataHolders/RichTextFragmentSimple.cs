using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders
{
    public class RichTextFragmentSimple
    {
        internal int Fragidx { get; set; }

        internal int OverallParagraphStartCharIdx { get; set; }
        //Char length
        internal int charStarIdxWithinCurrentLine { get; set; }
        //Width in points
        internal double Width { get; set; }
        //Font-size in points
        internal double FontSize { get; set; }

        public RichTextFragmentSimple() { }
    }
}
