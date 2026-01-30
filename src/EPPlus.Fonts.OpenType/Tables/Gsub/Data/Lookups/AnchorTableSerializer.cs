/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           AnchorTable serialization
 *************************************************************************************************/

namespace EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups
{
    /// <summary>
    /// Static helper for serializing AnchorTable structures.
    /// Used by MarkToBase, MarkToLigature, MarkToMark, and Cursive lookups.
    /// </summary>
    internal static class AnchorTableSerializer
    {
        /// <summary>
        /// Serializes an AnchorTable (Format 1, 2, or 3).
        /// </summary>
        /// <param name="writer">Binary writer</param>
        /// <param name="anchor">Anchor table to serialize</param>
        public static void Serialize(FontsBinaryWriter writer, AnchorTable anchor)
        {
            if (anchor == null)
                return;

            // Write format
            writer.WriteUInt16BigEndian(anchor.AnchorFormat);

            // Write X and Y coordinates (all formats)
            writer.WriteInt16BigEndian(anchor.XCoordinate);
            writer.WriteInt16BigEndian(anchor.YCoordinate);

            // Format-specific fields
            if (anchor.AnchorFormat == 2)
            {
                // Format 2: Anchor point index
                writer.WriteUInt16BigEndian(anchor.AnchorPoint);
            }
            else if (anchor.AnchorFormat == 3)
            {
                // Format 3: Device table offsets
                // We don't implement device tables yet, so write zeros
                writer.WriteUInt16BigEndian(0); // XDeviceOffset
                writer.WriteUInt16BigEndian(0); // YDeviceOffset
            }
        }
    }
}