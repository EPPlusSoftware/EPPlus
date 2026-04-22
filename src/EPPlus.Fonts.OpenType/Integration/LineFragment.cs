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
    [DebuggerDisplay("{Text}")]
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
        /// Index of the orginal TextFragment this line fragment came from
        /// </summary>
        public int FragmentIndex { get; set; }

        /// <summary>
        /// Width of a space in the original TextFragment
        /// </summary>
        public double SpaceWidth { get; internal set; }

        public string Text { get { return _finalFragmentText; } }

        //Seeing duplicated string in debugger creates clutter
        //Better to check Text and then realise the private variable is the actual value holder
        [DebuggerBrowsable(DebuggerBrowsableState.Never)]
        private string _finalFragmentText;

        /// <summary>
        /// Only to be set after finishing all operations
        /// Only to be set by TextLineSimple
        /// This class should be essentially read only after setting this
        /// </summary>
        /// <param name="text"></param>
        internal void SetFinalizedText(string text)
        {
            _finalFragmentText = text;
        }
        //public double AscentInPoints { get; private set; }
        //public double DescentInPoints { get; private set; }

        internal LineFragment(int rtFragmentIdx, int idxWithinLine/*, double ascentInPoints, double descentInPoints*/)
        {
            FragmentIndex = rtFragmentIdx;
            StartIdx = idxWithinLine;
            //AscentInPoints = ascentInPoints;
            //DescentInPoints = descentInPoints;
        }
    }
}
