using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class FeatureTable : FontTableElement
    {
        // USHORT FeatureParams (reserverad för framtiden, ska vara 0)
        public ushort FeatureParams { get; set; }

        // USHORT LookupCount: Antal Lookups som denna feature använder
        public ushort LookupCount { get; set; }

        // USHORT[] LookupListIndices: Index i den globala LookupListTable
        public ushort[] LookupListIndices { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 1. USHORT FeatureParams (Set to 0, per spec)
            writer.WriteUInt16BigEndian(this.FeatureParams);

            // 2. USHORT LookupListCount
            if (this.LookupListIndices == null)
            {
                writer.WriteUInt16BigEndian(0);
            }
            else
            {
                writer.WriteUInt16BigEndian((ushort)this.LookupListIndices.Length);
            }

            // 3. USHORT[] LookupListIndex
            if (this.LookupListIndices != null)
            {
                foreach (ushort lookupIndex in this.LookupListIndices)
                {
                    writer.WriteUInt16BigEndian(lookupIndex);
                }
            }
        }
    }
}
