using EPPlus.Fonts.OpenType.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    public class TextLineSimple
    {
        public List<LineFragment> LineFragments { get; internal set; } = new List<LineFragment>();

        public string Text { get; internal set; }
        /// <summary>
        /// The largest font size within this line
        /// </summary>
        public double LargestFontSize { get; internal set; }
        /// <summary>
        /// The largest ascent in points within this line 
        /// (Two differing fonts with the same size can have different ascents)
        /// </summary>
        public double LargestAscent { get; internal set; }
        /// <summary>
        /// The largest descent in points within this line 
        /// (Two differing fonts with the same size can have different descents)
        /// </summary>
        public double LargestDescent { get; internal set; }

        public double Width { get; internal set; }

        public double lastFontSpaceWidth { get; internal set; }

        internal bool WasWrappedOnSpace = false;

        /// <summary>
        /// In renderers like Excel the width of trailing spaces
        /// MUST be resepected in some cases.
        /// In others (e.g. Centering) the spaces width must be ignored
        /// </summary>
        public double GetWidthWithoutTrailingSpaces()
        {
            lastFontSpaceWidth = LineFragments.Last().SpaceWidth;

            var trailingSpaceCount = 0;

            for(int i = Text.Count()-1; i > 0; i--)
            {
                if(Text[i] != ' ')
                {
                    break;
                }

                trailingSpaceCount++;
            }

            if(WasWrappedOnSpace)
            {
                trailingSpaceCount++;
            }

            var widthWithoutTrail = Width - lastFontSpaceWidth * (trailingSpaceCount);
            return widthWithoutTrail;
        }

        public TextLineSimple()
        {
        }

        public TextLineSimple(string text, double largestFontSize, double largestAscent, double largestDescent)
        {
            Text = text;
            LargestFontSize = largestFontSize;
            LargestAscent = largestAscent;
            LargestDescent = largestDescent;
        }

        public string GetLineFragmentText(LineFragment rtFragment)
        {
            if (LineFragments.Contains(rtFragment) == false)
            {
                throw new InvalidOperationException($"GetFragmentText failed. Cannot retrieve {rtFragment} since it is not part of this textLine: {this}");
            }

            if (string.IsNullOrEmpty(Text))
            {
                return Text;
            }

            var startIdx = rtFragment.StartIdx;

            var idxInLst = LineFragments.FindIndex(x => x == rtFragment);
            if (idxInLst == LineFragments.Count - 1)
            {
                return Text.Substring(startIdx, Text.Length - startIdx);
            }
            else
            {
                var endIdx = LineFragments[idxInLst + 1].StartIdx;
                return Text.Substring(startIdx, endIdx - startIdx);
            }
        }

        internal LineFragment SplitAndGetLeftoverLineFragment(ref LineFragment origLf, double widthAtSplit)
        {
            //If we are splitting a fragment its position in the new line should be 0
            var newLineFragment = new LineFragment(origLf.RtFragIdx, 0);
            newLineFragment.Width = origLf.Width - widthAtSplit;

            origLf.Width = widthAtSplit;

            return newLineFragment;
        }
    }
}
