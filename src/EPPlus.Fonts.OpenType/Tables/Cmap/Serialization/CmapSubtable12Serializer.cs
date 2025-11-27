using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace EPPlus.Fonts.OpenType.Tables.Cmap.Serialization
{
    internal class CmapSubtable12Serializer : CmapSubtableSerializerBase<CmapSubtable12>
    {
        internal override void Serialize(CmapSubtable12 subTable, FontsBinaryWriter writer)
        {
            writer.WriteUInt16BigEndian(subTable.Format);
            writer.WriteUInt16BigEndian(subTable.Reserved);
            writer.WriteUInt32BigEndian(subTable.Length);
            writer.WriteUInt32BigEndian(subTable.Language);
            writer.WriteUInt32BigEndian(subTable.NumGroups);

            foreach (var group in subTable.Groups)
            {
                group.Serialize(writer);
            }
        }
    }
}
