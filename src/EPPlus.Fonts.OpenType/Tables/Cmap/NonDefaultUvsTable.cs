using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{

    public class NonDefaultUvsTable : FontTableElement
    {
        public List<UvsMapping> Mappings { get; internal set; } = new();

        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteUInt32BigEndian((uint)Mappings.Count);

            foreach (var mapping in Mappings)
            {
                mapping.Serialize(writer);
            }
        }
    }

}
