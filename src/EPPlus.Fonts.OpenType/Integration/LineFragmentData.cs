using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    [DebuggerDisplay("{Text}")]
    public class LineFragmentData
    {
        
        /// <summary>
        /// Width of this fragment
        /// </summary>
        public double Width { get { return _getWidth(); } }

        /// <summary>
        /// Text of this fragment
        /// </summary>
        public string Text { get; } 

        /// <summary>
        /// Original text fragment
        /// </summary>
        public TextFragment OriginalTextFragment { get { return _getTextFragment(); } }


        #region Callbacks
        Func<TextFragment> _getTextFragment;
        Func<double> _getWidth;
        #endregion

        internal LineFragmentData(Func<TextFragment> getTextFragment, Func<double> getWidth, string text)
        {
            _getTextFragment = getTextFragment;
            _getWidth = getWidth;
            Text = text;
        }
    }
}
