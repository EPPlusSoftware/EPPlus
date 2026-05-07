using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace EPPlus.Fonts.OpenType.Integration
{
    //Leaving this for now. May be neccesary for vertical text
    internal class CharInfo
    {
        /// <summary>
        /// Index within AllText
        /// </summary>
        internal int Index;
        /// <summary>
        /// The input Fragment Index
        /// </summary>
        internal int Fragment;
        /// <summary>
        /// The input char index within the fragment
        /// </summary>
        internal int InnerIndex;
        /// <summary>
        /// Width in points of the char
        /// </summary>
        internal int Width;


        internal int Line;

        internal bool IsSeparator { get; set; }

        public CharInfo(int index, int fragment, int fragCharIdx/*, int line*/) 
        {
            Index = index;
            Fragment = fragment;
            //Line = line;
        }
    }
}
