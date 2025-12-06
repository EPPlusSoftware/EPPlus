using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType
{
    public class FontSerializationContext
    {
        public OpenTypeFont Font { get; }
        public bool IsSubsetInProgress { get; } = false;

        public FontSerializationContext(OpenTypeFont font, bool isSubsetInProgress = false)
        {
            Font = font;
            IsSubsetInProgress = isSubsetInProgress;
        }
    }
}
