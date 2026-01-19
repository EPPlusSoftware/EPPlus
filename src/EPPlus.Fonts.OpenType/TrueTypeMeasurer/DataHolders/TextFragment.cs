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
        List<int> positionsToBreakAt = new();

        internal TextFragment(string thisFragment, int startPos)
        {
            ThisFragment = thisFragment;
            StartPos = startPos;
        }

        internal int GetEndPosition()
        {
            return StartPos + ThisFragment.Length; 
        }

        /// <summary>
        /// Line-breaks to be added after wrapping
        /// </summary>
        /// <param name="pos"></param>
        internal void AddLocalLbPositon(int pos)
        {
            var locaLnPos = pos - StartPos;
            positionsToBreakAt.Add(locaLnPos);
        }

        internal string GetWrappedFragment()
        {
            string alteredFragment = ThisFragment;

            for (int i = 0; i < positionsToBreakAt.Count; i++)
            {
                alteredFragment = alteredFragment.Insert(positionsToBreakAt[i] + i, Environment.NewLine);
            }

            return alteredFragment;
        }
    }
}
