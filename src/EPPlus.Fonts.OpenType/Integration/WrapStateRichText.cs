using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    internal class WrapStateRichText : WrapStateBase
    {
        internal LineFragment LineFrag = null;

        public WrapStateRichText(double lineWidth) 
        {
            CurrentLineWidth = lineWidth;
        }

        internal void EndCurrentTextLineAndIntializeNext(int fragIdx, int startIdxOfNewFragment)
        {
            CurrentTextLine.LineFragments.Add(LineFrag);
            Lines.Add(CurrentTextLine);

            CurrentTextLine = new TextLineSimple();
            LineFrag = new LineFragment(fragIdx, startIdxOfNewFragment);
        }
    }
}
