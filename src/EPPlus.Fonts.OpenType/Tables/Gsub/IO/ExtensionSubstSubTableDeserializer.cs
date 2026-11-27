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
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
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

            ushort format = _reader.ReadUInt16BigEndian();          // Extension format (must be 1)
            ushort extensionLookupType = _reader.ReadUInt16BigEndian(); // The wrapped substitution type
            uint offset = _reader.ReadUInt32BigEndian();             // 32-bit offset to the wrapped subtable

            if (format != 1) return null;

            long innerSubTableAbsoluteStart = absoluteStart + offset;
            FontTableElement inner;

            switch (extensionLookupType)
            {
                case 1:
                    inner = new SingleSubstSubTableDeserializer(_reader).Deserialize(innerSubTableAbsoluteStart);
                    break;
                case 2:
                    inner = new MultipleSubstSubTableDeserializer(_reader).Deserialize(innerSubTableAbsoluteStart);
                    break;
                case 4:
                    inner = new LigatureSubstSubTableDeserializer(_reader).Deserialize(innerSubTableAbsoluteStart);
                    break;
                case 6:
                    // Returns null for formats 1 and 2, which are not supported yet
                    inner = new ChainingContextualDeserializer(_reader).Deserialize(innerSubTableAbsoluteStart);
                    break;
                default:
                    // Types 3, 5 and 8 are not supported yet. Type 7 must never be wrapped in itself.
                    inner = null;
                    break;
            }

            if (inner == null) return null;

            // Keep the wrapper so the inner lookup type is preserved in the model, rather than
            // having to be guessed from the subtable's class later.
            return new ExtensionSubstSubTable
            {
                ExtensionLookupType = extensionLookupType,
                ExtendedSubTable = inner
            };
        }
    }
}