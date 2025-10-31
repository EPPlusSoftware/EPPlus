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
using EPPlus.Fonts.OpenType.Tables.Cmap.Serializers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    /// <summary>
    /// This is the standard character-to-glyph-index mapping subtable 
    /// for fonts that support only Unicode Basic Multilingual Plane characters 
    /// (U+0000 to U+FFFF).
    /// See https://docs.microsoft.com/en-us/typography/opentype/spec/cmap#format-4-segment-mapping-to-delta-values
    /// </summary>
    public class CmapSubtable4 : CmapSubtableBase
    {
        public CmapSubtable4(CmapTable parent)
        {
            _parentTable = parent;
        }

        private readonly CmapTable _parentTable;
        private Dictionary<ushort, char> _glyphIndextoCharMappings = new Dictionary<ushort, char>();
        private Dictionary<char, ushort> _CharMappingstoglyphIndex = new Dictionary<char, ushort>();
        private void OnMappingDone(char c, ushort gIx)
        {
            if(!_glyphIndextoCharMappings.ContainsKey(gIx))
            {
                _glyphIndextoCharMappings.Add(gIx, c);
            }
            if(!_CharMappingstoglyphIndex.ContainsKey(c))
            {
                _CharMappingstoglyphIndex.Add(c, gIx);
            }
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable4Serializer();
            serializer.Serialize(this, writer);
        }

        internal CmapSubtable4(FontsBinaryReader reader)
        {

            _reader = reader;
            _initialPos = reader.BaseStream.Position;

            Format = 4;
            Length = _reader.ReadUInt16BigEndian();
            Language = _reader.ReadUInt16BigEndian();
            SegCountX2 = _reader.ReadUInt16BigEndian();
            SearchRange = _reader.ReadUInt16BigEndian();
            EntrySelector = _reader.ReadUInt16BigEndian();
            RangeShift = _reader.ReadUInt16BigEndian();

            int segCount = SegCountX2 / 2;

            // Read segment arrays
            var endCodes = new ushort[segCount];
            for (int i = 0; i < segCount; i++)
            {
                endCodes[i] = _reader.ReadUInt16BigEndian();
            }

            ushort reservedPad = _reader.ReadUInt16BigEndian();

            var startCodes = new ushort[segCount];
            for (int i = 0; i < segCount; i++)
            {
                startCodes[i] = _reader.ReadUInt16BigEndian();
            }

            var idDeltas = new short[segCount];
            for (int i = 0; i < segCount; i++)
            {
                idDeltas[i] = _reader.ReadInt16BigEndian();
            }

            var idRangeOffsets = new ushort[segCount];
            for (int i = 0; i < segCount; i++)
            {
                idRangeOffsets[i] = _reader.ReadUInt16BigEndian();
            }

            // Calculate glyphIdArray start position
            long glyphArrayStart = _reader.BaseStream.Position;

            var segments = new List<CmapSubtable4Segment>();
            var glyphMappings = new List<GlyphMapping>();

            for (int i = 0; i < segCount; i++)
            {
                var segment = new CmapSubtable4Segment
                {
                    StartCode = startCodes[i],
                    EndCode = endCodes[i],
                    IdDelta = idDeltas[i],
                    IdRangeOffset = idRangeOffsets[i]
                };
                if (segment.StartCode == 0xFFFF && segment.EndCode == 0xFFFF)
                {
                    // Sentinel segment – skip mapping
                    continue;
                }

                if (segment.IdRangeOffset != 0)
                {
                    int rangeLength = segment.EndCode - segment.StartCode + 1;
                    long offset = glyphArrayStart + (2 * (i + segment.IdRangeOffset / 2));
                    long currentPos = _reader.BaseStream.Position;

                    _reader.BaseStream.Seek(offset, SeekOrigin.Begin);
                    segment.GlyphIdArray = new ushort[rangeLength];
                    for (int j = 0; j < rangeLength; j++)
                    {
                        segment.GlyphIdArray[j] = _reader.ReadUInt16BigEndian();
                        ushort glyphIndex = segment.GlyphIdArray[j];
                        ushort charCode = (ushort)(segment.StartCode + j);
                        if (glyphIndex != 0)
                        {
                            glyphMappings.Add(new GlyphMapping
                            {
                                CharacterCode = charCode,
                                GlyphIndex = glyphIndex
                            });
                            OnMappingDone((char)charCode, glyphIndex);
                        }
                    }

                    _reader.BaseStream.Seek(currentPos, SeekOrigin.Begin);
                }
                else
                {
                    for (ushort c = segment.StartCode; c <= segment.EndCode; c++)
                    {
                        ushort glyphIndex = (ushort)((c + segment.IdDelta) % 65536);
                        if (glyphIndex != 0)
                        {
                            glyphMappings.Add(new GlyphMapping
                            {
                                CharacterCode = c,
                                GlyphIndex = glyphIndex
                            });
                            OnMappingDone((char)c, glyphIndex);
                        }
                    }
                }

                segments.Add(segment);
            }

            Segments = segments;
            GlyphMappingArray = glyphMappings.ToArray();
        }

        private readonly FontsBinaryReader _reader;
        private readonly long _initialPos;

        public ushort SegCountX2 { get; set; }

        public ushort SearchRange { get; private set; }

        public ushort EntrySelector { get; private set; }

        public ushort RangeShift { get; private set; }

        public IDictionary<ushort, char> GlyphIndexToCharMappings => _glyphIndextoCharMappings;
        public IDictionary<char, ushort> CharMappingsToGlyphIndex => _CharMappingstoglyphIndex;

        public override ushort Format { get; }

        public override ushort Length { get; }

        public override ushort Language { get; }

        public override GlyphMapping[] GlyphMappingArray { get; }

        internal List<CmapSubtable4Segment> Segments { get; } = new();

        /// <summary>
        /// Returns glyph index of a character
        /// </summary>
        /// <param name="c"></param>
        /// <returns>The glyph index or 0 if the glyph doesn't exist</returns>
        public ushort GetGlyphIndex(char c)
        {
            if(_CharMappingstoglyphIndex != null && _CharMappingstoglyphIndex.TryGetValue(c, out ushort gi))
            {
                return gi;
            }
            return 0;
        }
    }
}
