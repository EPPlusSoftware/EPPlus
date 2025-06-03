using System;
using System.Collections.Generic;

namespace FontLab1.Tables.Cmap
{
    /// <summary>
    /// This is the standard character-to-glyph-index mapping subtable 
    /// for fonts that support only Unicode Basic Multilingual Plane characters 
    /// (U+0000 to U+FFFF).
    /// See https://docs.microsoft.com/en-us/typography/opentype/spec/cmap#format-4-segment-mapping-to-delta-values
    /// </summary>
    internal class CmapSubtable4
    {
        private Dictionary<ushort, char> _glyphIndextoCharMappings = new Dictionary<ushort, char>();
        private void OnMappingDone(char c, ushort gIx)
        {
            if(!_glyphIndextoCharMappings.ContainsKey(gIx))
            {
                _glyphIndextoCharMappings.Add(gIx, c);
            }
        }

        public CmapSubtable4(MyBinaryReader reader)
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

            // start reading items...
            var segCount = SegCountX2 / 2;
            var endCodes = new List<ushort>();
            for(var x = 0; x < segCount; x++)
            {
                endCodes.Add(_reader.ReadUInt16BigEndian());
            }
            var reservedPad = reader.ReadUInt16BigEndian();
            var startCodes = new List<ushort>();
            for (var x = 0; x < segCount; x++)
            {
                startCodes.Add(_reader.ReadUInt16BigEndian());
            }
            var idDeltas = new List<short>();
            for (var x = 0; x < segCount; x++)
            {
                idDeltas.Add(_reader.ReadInt16BigEndian());
            }
            var idRangeOffsets = new List<ushort>();
            for (var x = 0; x < segCount; x++)
            {
                idRangeOffsets.Add(_reader.ReadUInt16BigEndian());
            }
            var glyphMappings = new List<GlyphMapping>();
            var glyphIds = default(ushort[]);
            for(var seg = 0; seg < segCount - 1; seg++)
            {
                var rangeOffset = idRangeOffsets[seg];
                var delta = idDeltas[seg];
                for(var cc = startCodes[seg]; cc < endCodes[seg]; cc++)
                {
                    if(rangeOffset == 0)
                    {
                        glyphMappings.Add(new GlyphMapping
                        {
                            CharacterCode = cc,
                            GlyphIndex = (ushort)(delta + cc)
                        });
                        OnMappingDone(Convert.ToChar(cc), (ushort)(delta + cc));
                    }
                    else
                    {
                        if(glyphIds == null)
                        {
                            var gIds = new List<ushort>();
                            var hLength = reader.BaseStream.Position - _initialPos + 2;
                            var nGlyphs = (Length - hLength) / 2;
                            for(var x = 0; x < nGlyphs; x++)
                            {
                                var gId = _reader.ReadUInt16BigEndian();
                                gIds.Add(gId);
                            }
                            glyphIds = gIds.ToArray();
                        }
                        long offset = (rangeOffset / 2) + (cc - startCodes[seg]);
                        var arrayIndex = offset - startCodes.Count + seg;
                        var gm = new GlyphMapping
                        {
                            CharacterCode = cc,
                            GlyphIndex = glyphIds[arrayIndex]
                        };
                        glyphMappings.Add(gm);
                        OnMappingDone(gm.Char, gm.GlyphIndex);
                    }
                }
            }
            GlyphMappingArray = glyphMappings.ToArray();
        }

        private readonly MyBinaryReader _reader;
        private readonly long _initialPos;

        public ushort Format { get; set; }

        public ushort Length { get; set; }

        public ushort Language { get; set; }

        public ushort SegCountX2 { get; set; }

        public ushort SearchRange { get; private set; }

        public ushort EntrySelector { get; private set; }

        public ushort RangeShift { get; private set; }

        public GlyphMapping[] GlyphMappingArray { get; set; }

        public IDictionary<ushort, char> GlyphIndexToCharMappings => _glyphIndextoCharMappings;
    }
}
