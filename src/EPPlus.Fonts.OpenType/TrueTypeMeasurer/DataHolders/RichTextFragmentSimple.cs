using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders
{
    public class RichTextFragmentSimple
    {
        public int Fragidx { get; internal set; }

        internal int OverallParagraphStartCharIdx { get; set; }
        //Char length
        internal int charStarIdxWithinCurrentLine { get; set; }
        /// <summary>
        /// Width in points
        /// </summary>
        public double Width { get; internal set; }

        public RichTextFragmentSimple() { }

        internal RichTextFragmentSimple Clone()
        {
            var fragment = new RichTextFragmentSimple();
            fragment.Width = Width;
            fragment.OverallParagraphStartCharIdx = OverallParagraphStartCharIdx;
            fragment.charStarIdxWithinCurrentLine = charStarIdxWithinCurrentLine;
            fragment.Fragidx = Fragidx;
            return fragment;
        }
    }
}
