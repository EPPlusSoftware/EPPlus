using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    internal class WrapStateRichText : WrapStateBase
    {
        public WrapStateRichText(double lineWidth) 
        {
            CurrentLineWidth = lineWidth;
        }
    }
}
