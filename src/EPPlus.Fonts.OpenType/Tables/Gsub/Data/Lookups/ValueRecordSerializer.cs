/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           ValueRecord serialization
 *************************************************************************************************/

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups
{
    /// <summary>
    /// Static helper for serializing ValueRecord structures.
    /// Used by SinglePos, PairPos, and other positioning lookups.
    /// </summary>
    internal static class ValueRecordSerializer
    {
        /// <summary>
        /// Serializes a ValueRecord based on the ValueFormat flags.
        /// ValueFormat is a bit field indicating which fields are present.
        /// </summary>
        /// <param name="writer">Binary writer</param>
        /// <param name="record">ValueRecord to serialize</param>
        /// <param name="valueFormat">Bit flags indicating which fields to write</param>
        public static void Serialize(FontsBinaryWriter writer, ValueRecord record, ushort valueFormat)
        {
            if (record == null)
                return;

            // Bit 0x0001: XPlacement
            if ((valueFormat & 0x0001) != 0)
                writer.WriteInt16BigEndian(record.XPlacement);

            // Bit 0x0002: YPlacement
            if ((valueFormat & 0x0002) != 0)
                writer.WriteInt16BigEndian(record.YPlacement);

            // Bit 0x0004: XAdvance
            if ((valueFormat & 0x0004) != 0)
                writer.WriteInt16BigEndian(record.XAdvance);

            // Bit 0x0008: YAdvance
            if ((valueFormat & 0x0008) != 0)
                writer.WriteInt16BigEndian(record.YAdvance);

            // Bits 0x0010-0x0080: Device tables (not implemented yet)
            // XPlaDevice, YPlaDevice, XAdvDevice, YAdvDevice
            // We skip these for now
        }
    }
}