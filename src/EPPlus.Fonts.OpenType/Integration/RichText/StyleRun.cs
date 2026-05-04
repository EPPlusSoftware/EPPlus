using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration.RichText
{
    internal class StyleRun : TextSection
    {
        internal int FragmentIndex { get; private set; }

        internal StyleRun(int fragmentIndex, int startIdx, int endIndex, Func<string> getFullText, Func<int, int, string> getText) : base(startIdx, endIndex, getFullText, getText)
        {
            FragmentIndex = fragmentIndex;
        }

        private double[] _charWidths;

        internal void SetCharWidths(double[] charWidths)
        {
            _charWidths = charWidths;
        }
    }
}
