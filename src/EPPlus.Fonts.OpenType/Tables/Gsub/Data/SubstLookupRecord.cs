using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data
{
    public class SubstLookupRecord
    {
        public ushort SequenceIndex { get; set; } // Vilken glyf i Input-sekvensen?
        public ushort LookupListIndex { get; set; } // Vilken Lookup ska köras?
    }
}
