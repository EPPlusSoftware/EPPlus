using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class FeatureRecord : FontTableElement
    {
        public Tag FeatureTag { get; set; }
        // Offset till FeatureTable, relativt till FeatureListTable start
        public ushort FeatureOffset { get; set; }

        // Den faktiska feature-tabellen
        public FeatureTable FeatureTable { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            throw new NotImplementedException();
        }
    }
}
