using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer
{
    internal class CharInfo
    {
        internal int Index;
        internal int Fragment;
        internal int Line;

        internal int Width;

        public CharInfo(int index, int fragment, int line) 
        {
            Index = index;
            Fragment = fragment;
            Line = line;
        }
    }
}
