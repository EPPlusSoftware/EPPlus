using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class VariationSelector : FontTableElement
    {
        public uint VarSelector { get; internal set; } // 24-bit värde
        public uint DefaultUVSOffset { get; internal set; }
        public uint NonDefaultUVSOffset { get; internal set; }

        public DefaultUvsTable DefaultUvsTable { get; internal set; }
        public NonDefaultUvsTable NonDefaultUvsTable { get; internal set; }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteUInt24BigEndian(VarSelector);
            writer.WriteUInt32BigEndian(DefaultUVSOffset);
            writer.WriteUInt32BigEndian(NonDefaultUVSOffset);
        }
    }
}
