using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.TextShaping.DataHolders
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
        internal List<double> PointWidthPerOutputLineInFragment = new();
        internal List<double> PointYIncreasePerLine = new();

        List<int> positionsToBreakAt = new();

        internal float FontSize { get; set; } = float.NaN;

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

        internal string GetLineBreakFreeFragment()
        {
            string freeFragment = ThisFragment.Replace("\r", "");
            freeFragment.Replace("\n", "");
            return freeFragment;
        }

        internal string GetWrappedFragment()
        {
            if(positionsToBreakAt.Count == 0)
            {
                return ThisFragment;
            }

            string alteredFragment = ThisFragment;
            int deleteSpaceCount = 0;
            for (int i = 0; i < positionsToBreakAt.Count; i++)
            {
                var insertPosition = positionsToBreakAt[i] + i - deleteSpaceCount;
                alteredFragment = alteredFragment.Insert(insertPosition, Environment.NewLine);
                //Spaces after inserted newlines are to be removed
                if (alteredFragment[insertPosition + Environment.NewLine.Length] == ' ')
                {
                    alteredFragment = alteredFragment.Remove(insertPosition + Environment.NewLine.Length, 1);
                    deleteSpaceCount++;
                }
            }

            return alteredFragment;
        }
    }
}
