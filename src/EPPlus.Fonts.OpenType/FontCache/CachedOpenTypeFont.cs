using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.FontCache
{
    internal class CachedOpenTypeFont
    {
        public OpenTypeFont Font { get; set; }

        public bool IsLoaded { get; set; }
    }
}
