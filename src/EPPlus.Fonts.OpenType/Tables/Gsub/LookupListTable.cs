using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class LookupListTable : FontTableElement
    {
        public List<LookupTable> Lookups { get; set; } = new List<LookupTable>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            throw new NotImplementedException();
        }
    }
}
