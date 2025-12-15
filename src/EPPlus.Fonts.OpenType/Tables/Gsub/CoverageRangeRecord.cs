using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class CoverageRangeRecord : FontTableElement
    {
        public ushort StartGlyphID { get; set; }
        public ushort EndGlyphID { get; set; }
        public ushort StartCoverageIndex { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            throw new NotImplementedException();
        }
    }
}
