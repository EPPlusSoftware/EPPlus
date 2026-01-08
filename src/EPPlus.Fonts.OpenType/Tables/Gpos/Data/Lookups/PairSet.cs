/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/07/2026         EPPlus Software AB           GPOS PairSet
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups
{
    /// <summary>
    /// PairSet - list of pairs for a specific first glyph
    /// </summary>
    public class PairSet
    {
        public List<PairValueRecord> PairValueRecords { get; set; }

        internal void Serialize(FontsBinaryWriter writer, ushort valueFormat1, ushort valueFormat2)
        {
            writer.WriteUInt16BigEndian((ushort)(PairValueRecords?.Count ?? 0));

            if (PairValueRecords != null)
            {
                foreach (var record in PairValueRecords)
                {
                    record.Serialize(writer, valueFormat1, valueFormat2);
                }
            }
        }
    }
}
