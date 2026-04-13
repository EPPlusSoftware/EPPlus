using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    public class TextLine
    {
        internal List<int> richTextIndicies;
        internal string content;
        internal int startIndex;
        internal List<int> rtContentStartIndexPerRt;
        internal int lastRtInternalIndex;
        internal int startRtInternalIndex;
    }
}
