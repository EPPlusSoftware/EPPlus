using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// Finalized output which uses callbacks to get data but can never have data set.
    /// </summary>
    [DebuggerDisplay("{Text}")]
    public class LineFragmentOutput
    {
        /// <summary>
        /// Char idx within the whole input
        /// </summary>
        public int IdxWithinOriginal { get { return _getStartIdxOrig(); } }
        /// <summary>
        /// Char idx within the line
        /// </summary>
        public int StartIdx { get { return _getStartIdx(); } }
        /// <summary>
        /// Width of this fragment
        /// </summary>
        public double Width { get { return _getWidth(); } }

        /// <summary>
        /// Text of this fragment
        /// Is not a callback since the LineFragment class does not store the string data directly.
        /// This is for performance reasons.
        /// </summary>
        public string Text { get; } 

        /// <summary>
        /// Original text fragment
        /// If this becomes null/loses its reference the original textfragment has been deleted
        /// Don't delete/clear the original textFragment list if you plan to use/check it here.
        /// </summary>
        public TextFragment OriginalTextFragment { get { return _getTextFragment(); } }


        /// <summary>
        /// Looks up the data in the internal class LineFragment instead of making copies
        /// </summary>
        #region Callbacks
        Func<TextFragment> _getTextFragment;
        Func<double> _getWidth;
        Func<int> _getStartIdx;
        Func<int> _getStartIdxOrig;
        #endregion

        internal LineFragmentOutput(Func<TextFragment> getTextFragment, Func<double> getWidth, Func<int> startIdx, Func<int> startOrig, string text)
        {
            _getTextFragment = getTextFragment;
            _getWidth = getWidth;
            _getStartIdx = startIdx;
            _getStartIdxOrig = startOrig;
            Text = text;
        }
    }
}
