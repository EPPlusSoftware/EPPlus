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
using System.IO;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.IO
{
    /// <summary>
    /// Deserializes Extension Substitution subtables (Lookup Type 7).
    /// These are used to provide 32-bit offsets to other substitution types when the 16-bit 
    /// limit of the standard tables is exceeded.
    /// </summary>
    internal class ExtensionSubstSubTableDeserializer
    {
        private readonly FontsBinaryReader _reader;

        public ExtensionSubstSubTableDeserializer(FontsBinaryReader reader)
        {
            _reader = reader;
        }

        public FontTableElement Deserialize(long absoluteStart)
        {
            _reader.BaseStream.Seek(absoluteStart, SeekOrigin.Begin);

            ushort format = _reader.ReadUInt16BigEndian(); // Extension format (must be 1)
            ushort lookuptype = _reader.ReadUInt16BigEndian(); // The actual substitution type being extended
            uint offset = _reader.ReadUInt32BigEndian(); // 32-bit offset to the actual subtable

            long innerSubTableAbsoluteStart = absoluteStart + offset;

            // Redirect to the appropriate deserializer based on the encapsulated lookup type
            switch (lookuptype)
            {
                case 1:
                    return new SingleSubstSubTableDeserializer(_reader).Deserialize(innerSubTableAbsoluteStart);
                case 4:
                    return new LigatureSubstSubTableDeserializer(_reader).Deserialize(innerSubTableAbsoluteStart);
                case 6:
                    // Note: We will add the ChainingContextualDeserializer here once implemented
                    return null;
                default:
                    // Other types can be added here as support grows
                    return null;
            }
        }
    }
}