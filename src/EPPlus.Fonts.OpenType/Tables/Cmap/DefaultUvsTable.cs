using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class DefaultUvsTable : FontTableElement
    {
        public uint NumUnicodeValueRanges { get; internal set; }

        public List<UnicodeRange> Ranges { get; internal set; } = new List<UnicodeRange>();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteUInt32BigEndian((uint)Ranges.Count);

            foreach (var range in Ranges)
            {
                range.Serialize(writer);
            }
        }
    }
}
