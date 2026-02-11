using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// The fragment of a richTextFragment that is within a line
    /// </summary>
    public class LineFragment
    {
        public int StartIdx { get; set; }
        public double Width { get; set; }
        public int RtFragIdx { get; set; }

        internal LineFragment(int rtFragmentIdx, int idxWithinLine)
        {
            RtFragIdx = rtFragmentIdx;
            StartIdx = idxWithinLine;
        }
    }
}
