using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders
{
    public class TextLineSimple
    {
        public List<RichTextFragmentSimple> RtFragments { get; private set; } = new List<RichTextFragmentSimple>();

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

        /// <summary>
        /// Get the text of one of the fragments in this line
        /// </summary>
        /// <param name=""></param>
        /// <returns></returns>
        public string GetFragmentText(RichTextFragmentSimple rtFragment)
        {
            if(RtFragments.Contains(rtFragment) == false)
            {
                throw new InvalidOperationException($"GetFragmentText failed. Cannot retrieve {rtFragment} since it is not part of this textLine: {this}");
            }

            var startIdx = rtFragment.charStarIdxWithinCurrentLine;

            var idxInLst = RtFragments.FindIndex(x => x ==  rtFragment);
            if(idxInLst == RtFragments.Count -1)
            {
                return Text.Substring(startIdx, Text.Length - startIdx);
            }
            else
            {
                var endIdx = RtFragments[idxInLst + 1].charStarIdxWithinCurrentLine;
                return Text.Substring(startIdx, endIdx - startIdx);
            }
        }
    }
}
