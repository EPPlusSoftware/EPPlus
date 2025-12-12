/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
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
