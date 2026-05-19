using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration.RichText
{
    public class StyleRun : TextSection
    {
        internal int FragmentIndex { get; private set; }
        internal double SpaceWidth { get; private set; }

        /// <summary>
        /// A styleRun or "Span" class
        /// </summary>
        /// <param name="fragmentIndex"></param>
        /// <param name="startIdx"></param>
        /// <param name="endIndex"></param>
        /// <param name="getFullText"></param>
        /// <param name="getText"></param>
        internal StyleRun(int fragmentIndex, int startIdx, int endIndex, Func<string> getFullText, Func<int, int, string> getText) : base(startIdx, endIndex, getFullText, getText)
        {
            FragmentIndex = fragmentIndex;
        }

        private double[] _charWidths;

        internal void SetCharWidths(double[] charWidths, double spaceWidth)
        {
            _charWidths = charWidths;
            SpaceWidth = spaceWidth;
        }

        internal double GetCharWidthByIndex(int index)
        {
            return _charWidths[index];
        }
    }
}
