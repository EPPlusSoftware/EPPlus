using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    public interface IFontSubsetProcessor
    {
        void Process(FontSubsettingContext context);
    }
}
