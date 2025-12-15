using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class LangSysTable : FontTableElement
    {
        // USHORT LookupOrder (Reserverat, ska vara 0)
        public ushort LookupOrder { get; set; }

        // USHORT RequiredFeatureIndex (Index till FeatureList. 0xFFFF om inget krävs)
        public ushort RequiredFeatureIndex { get; set; }

        // USHORT FeatureIndexCount
        public ushort FeatureIndexCount { get; set; }

        // USHORT[] FeatureIndices (Index till FeatureList)
        public ushort[] FeatureIndices { get; set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 1. USHORT LookupOrder
            writer.WriteUInt16BigEndian(this.LookupOrder);

            // 2. USHORT RequiredFeatureIndex
            writer.WriteUInt16BigEndian(this.RequiredFeatureIndex);

            // 3. USHORT FeatureIndexCount
            ushort count = this.FeatureIndices != null ? (ushort)this.FeatureIndices.Length : (ushort)0;
            writer.WriteUInt16BigEndian(count);

            // 4. USHORT[] FeatureIndices
            if (this.FeatureIndices != null)
            {
                foreach (ushort index in this.FeatureIndices)
                {
                    writer.WriteUInt16BigEndian(index);
                }
            }
        }
    }
}
