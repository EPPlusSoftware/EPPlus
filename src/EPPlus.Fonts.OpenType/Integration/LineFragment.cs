using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// The fragment of a richTextFragment that is within a line
    /// </summary>
    public class LineFragment
    {
        /// <summary>
        /// Start char position within TextLineSimple.Text
        /// </summary>
        public int StartIdx { get; set; }
        /// <summary>
        /// Width of this fragment
        /// </summary>
        public double Width { get; set; }
        /// <summary>
        /// Index of original TextFragment
        /// </summary>
        public int FragmentIndex { get; set; }

        /// <summary>
        /// Width of a space in the original TextFragment
        /// </summary>
        public double SpaceWidth { get; internal set; }

        internal LineFragment(int rtFragmentIdx, int idxWithinLine)
        {
            FragmentIndex = rtFragmentIdx;
            StartIdx = idxWithinLine;
        }
    }
}
