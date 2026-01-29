using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders
{
    public class TextLineSimple
    {
        internal List<RichTextFragmentSimple> RtFragments = new();

        public string Text { get; internal set; }
        internal double LargestFontSize { get; set; }
        internal double LargestAscent { get; set; }
        internal double LargestDescent { get; set; }
        internal double Width { get; set; }

        public TextLineSimple()
        {

        }

        public TextLineSimple(List<RichTextFragmentSimple> rtFragments, string text, double largestFontSize, double largestAscent, double largestDescent)
        {
            RtFragments = rtFragments;
            Text = text;
            LargestFontSize = largestFontSize;
            LargestAscent = largestAscent;
            LargestDescent = largestDescent;
        }
    }
}
