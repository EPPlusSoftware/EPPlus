/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/12/2026         EPPlus Software AB           Helper for layout table serialization
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables;
using System.IO;

namespace EPPlus.Fonts.OpenType.Utils
{
    /// <summary>
    /// Helper methods for serializing OpenType layout tables (GSUB, GPOS, GDEF, etc.).
    /// </summary>
    internal static class LayoutTableSerializationHelper
    {
        /// <summary>
        /// Updates an offset placeholder and serializes an element.
        /// Common pattern used by layout tables to write sub-tables.
        /// </summary>
        /// <param name="writer">Binary writer</param>
        /// <param name="tableStart">Start position of the main table</param>
        /// <param name="placeholderPos">Position where offset placeholder was written</param>
        /// <param name="element">Element to serialize</param>
        public static void UpdateOffsetAndSerialize(FontsBinaryWriter writer, long tableStart, long placeholderPos, FontTableElement element)
        {
            ushort offset = (ushort)(writer.BaseStream.Position - tableStart);
            long resumePos = writer.BaseStream.Position;

            // Update placeholder with calculated offset
            writer.BaseStream.Seek(placeholderPos, SeekOrigin.Begin);
            writer.WriteUInt16BigEndian(offset);

            // Resume and serialize element
            writer.BaseStream.Seek(resumePos, SeekOrigin.Begin);
            element.Serialize(writer);
        }
    }
}