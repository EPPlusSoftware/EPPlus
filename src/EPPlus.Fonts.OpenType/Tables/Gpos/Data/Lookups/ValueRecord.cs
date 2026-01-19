/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/07/2026         EPPlus Software AB           GPOS ValueRecord implementation
 *************************************************************************************************/
using System;

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups
{
    /// <summary>
    /// Represents a GPOS ValueRecord which defines positioning adjustments.
    /// Fields present are determined by ValueFormat flags.
    /// </summary>
    public class ValueRecord
    {
        /// <summary>
        /// Horizontal adjustment for placement (in font design units)
        /// </summary>
        public short XPlacement { get; set; }

        /// <summary>
        /// Vertical adjustment for placement (in font design units)
        /// </summary>
        public short YPlacement { get; set; }

        /// <summary>
        /// Horizontal adjustment for advance width (in font design units)
        /// Used for kerning!
        /// </summary>
        public short XAdvance { get; set; }

        /// <summary>
        /// Vertical adjustment for advance height (in font design units)
        /// </summary>
        public short YAdvance { get; set; }

        /// <summary>
        /// Device table for XPlacement (not implemented yet)
        /// </summary>
        public ushort XPlacementDeviceOffset { get; set; }

        /// <summary>
        /// Device table for YPlacement (not implemented yet)
        /// </summary>
        public ushort YPlacementDeviceOffset { get; set; }

        /// <summary>
        /// Device table for XAdvance (not implemented yet)
        /// </summary>
        public ushort XAdvanceDeviceOffset { get; set; }

        /// <summary>
        /// Device table for YAdvance (not implemented yet)
        /// </summary>
        public ushort YAdvanceDeviceOffset { get; set; }

        /// <summary>
        /// Reads a ValueRecord from reader based on format flags
        /// </summary>
        internal static ValueRecord Read(FontsBinaryReader reader, ushort valueFormat)
        {
            var record = new ValueRecord();

            if ((valueFormat & (ushort)ValueFormatFlags.XPlacement) != 0)
                record.XPlacement = reader.ReadInt16BigEndian();

            if ((valueFormat & (ushort)ValueFormatFlags.YPlacement) != 0)
                record.YPlacement = reader.ReadInt16BigEndian();

            if ((valueFormat & (ushort)ValueFormatFlags.XAdvance) != 0)
                record.XAdvance = reader.ReadInt16BigEndian();

            if ((valueFormat & (ushort)ValueFormatFlags.YAdvance) != 0)
                record.YAdvance = reader.ReadInt16BigEndian();

            if ((valueFormat & (ushort)ValueFormatFlags.XPlacementDevice) != 0)
                record.XPlacementDeviceOffset = reader.ReadUInt16BigEndian();

            if ((valueFormat & (ushort)ValueFormatFlags.YPlacementDevice) != 0)
                record.YPlacementDeviceOffset = reader.ReadUInt16BigEndian();

            if ((valueFormat & (ushort)ValueFormatFlags.XAdvanceDevice) != 0)
                record.XAdvanceDeviceOffset = reader.ReadUInt16BigEndian();

            if ((valueFormat & (ushort)ValueFormatFlags.YAdvanceDevice) != 0)
                record.YAdvanceDeviceOffset = reader.ReadUInt16BigEndian();

            return record;
        }

        /// <summary>
        /// Writes ValueRecord based on format flags
        /// </summary>
        internal void Write(FontsBinaryWriter writer, ushort valueFormat)
        {
            if ((valueFormat & (ushort)ValueFormatFlags.XPlacement) != 0)
                writer.WriteInt16BigEndian(XPlacement);

            if ((valueFormat & (ushort)ValueFormatFlags.YPlacement) != 0)
                writer.WriteInt16BigEndian(YPlacement);

            if ((valueFormat & (ushort)ValueFormatFlags.XAdvance) != 0)
                writer.WriteInt16BigEndian(XAdvance);

            if ((valueFormat & (ushort)ValueFormatFlags.YAdvance) != 0)
                writer.WriteInt16BigEndian(YAdvance);

            if ((valueFormat & (ushort)ValueFormatFlags.XPlacementDevice) != 0)
                writer.WriteUInt16BigEndian(XPlacementDeviceOffset);

            if ((valueFormat & (ushort)ValueFormatFlags.YPlacementDevice) != 0)
                writer.WriteUInt16BigEndian(YPlacementDeviceOffset);

            if ((valueFormat & (ushort)ValueFormatFlags.XAdvanceDevice) != 0)
                writer.WriteUInt16BigEndian(XAdvanceDeviceOffset);

            if ((valueFormat & (ushort)ValueFormatFlags.YAdvanceDevice) != 0)
                writer.WriteUInt16BigEndian(YAdvanceDeviceOffset);
        }

        /// <summary>
        /// Calculates ValueFormat flags needed for this record
        /// </summary>
        public ushort GetRequiredFormat()
        {
            ushort format = 0;

            if (XPlacement != 0)
                format |= (ushort)ValueFormatFlags.XPlacement;
            if (YPlacement != 0)
                format |= (ushort)ValueFormatFlags.YPlacement;
            if (XAdvance != 0)
                format |= (ushort)ValueFormatFlags.XAdvance;
            if (YAdvance != 0)
                format |= (ushort)ValueFormatFlags.YAdvance;
            if (XPlacementDeviceOffset != 0)
                format |= (ushort)ValueFormatFlags.XPlacementDevice;
            if (YPlacementDeviceOffset != 0)
                format |= (ushort)ValueFormatFlags.YPlacementDevice;
            if (XAdvanceDeviceOffset != 0)
                format |= (ushort)ValueFormatFlags.XAdvanceDevice;
            if (YAdvanceDeviceOffset != 0)
                format |= (ushort)ValueFormatFlags.YAdvanceDevice;

            return format;
        }

        /// <summary>
        /// Checks if this record is empty (all zeros)
        /// </summary>
        public bool IsEmpty()
        {
            return XPlacement == 0 && YPlacement == 0 &&
                   XAdvance == 0 && YAdvance == 0 &&
                   XPlacementDeviceOffset == 0 && YPlacementDeviceOffset == 0 &&
                   XAdvanceDeviceOffset == 0 && YAdvanceDeviceOffset == 0;
        }
    }

    /// <summary>
    /// Flags indicating which fields are present in a ValueRecord
    /// </summary>
    [Flags]
    public enum ValueFormatFlags : ushort
    {
        None = 0x0000,
        XPlacement = 0x0001,
        YPlacement = 0x0002,
        XAdvance = 0x0004,
        YAdvance = 0x0008,
        XPlacementDevice = 0x0010,
        YPlacementDevice = 0x0020,
        XAdvanceDevice = 0x0040,
        YAdvanceDevice = 0x0080
        // Reserved: 0xFF00
    }
}