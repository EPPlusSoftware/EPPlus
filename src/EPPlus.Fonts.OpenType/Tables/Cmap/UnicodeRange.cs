using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class UnicodeRange : FontTableElement
    {
        public uint StartUnicodeValue { get; internal set; }
        public byte AdditionalCount { get; internal set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteUInt24BigEndian(StartUnicodeValue);
            writer.Write(AdditionalCount);
        }
    }
}
