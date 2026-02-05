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
namespace EPPlus.Fonts.OpenType.Tables.Common.Layout.ClassDef
{
    /// <summary>
    /// Base class for ClassDef subtables (Format 1 and Format 2).
    /// Defines the common API and serialization entry point.
    /// </summary>
    public abstract class ClassDefTable : FontTableElement
    {
        /// <summary>
        /// ClassDef format (1 or 2).
        /// </summary>
        public ushort Format { get; protected set; }

        /// <summary>
        /// Returns the class value for a given glyph ID.
        /// </summary>
        public abstract int GetClass(ushort glyphId);

        /// <summary>
        /// Writes the ClassDef table to the stream.
        /// Subclasses implement the actual body.
        /// </summary>
        internal override void Serialize(FontsBinaryWriter writer)
        {
            writer.WriteUInt16BigEndian(Format);
            SerializeBody(writer);
        }

        /// <summary>
        /// Subclasses implement the format-specific serialization.
        /// </summary>
        internal abstract void SerializeBody(FontsBinaryWriter writer);
    }
}
