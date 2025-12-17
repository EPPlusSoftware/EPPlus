using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer
{
    //TextFragment is a portion of a longer string
    //the fragment itself may contain line-breaks
    internal class TextFragment
    {
        /// <summary>
        /// Char index starting position for this fragment
        /// </summary>
        internal int StartPos;

        internal List<int> IsPartOfLines;
        internal string ThisFragment;

        //string[] Lines;

        internal TextFragment(string thisFragment, int startPos)
        {
            ThisFragment = thisFragment;
            StartPos = startPos;

            //Lines = thisFragment.Split([Environment.NewLine], StringSplitOptions.None);
        }

        internal int GetEndPosition()
        {
            return StartPos + ThisFragment.Length; 
        }
    }
}
