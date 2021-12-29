/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Glyph
{
    internal class GlyphTableLoader : TableLoader<GlyphTable>
    {
        public GlyphTableLoader(BigEndianBinaryReader reader, Dictionary<string, TableRecord> tables) : base(reader, tables, TableNames.Glyf)
        {
            _glyphIndexes = TableLoaders.GetLocaTableLoader(reader, tables).Load().Offsets;
            _emptyGlyph = TableLoaders.GetHeadTableLoader(reader, tables).Load().GetDefaultBounds();
        }

        private readonly uint[] _glyphIndexes;
        private readonly BoundingRectangle _emptyGlyph;

        protected override GlyphTable LoadInternal()
        {
            var glyphHeaders = new GlyphHeader[_glyphIndexes.Length];
            for (var x = 0; x < _glyphIndexes.Length - 1; x++)
            {
                var ix = _glyphIndexes[x];
                _reader.BaseStream.Position = _offset + ix;
                if (ix == _glyphIndexes[x + 1])
                {
                    glyphHeaders[x] = new GlyphHeader(0, _emptyGlyph);
                    continue;
                }
                var numberOfContours = _reader.ReadInt16BigEndian();
                var xMin = _reader.ReadInt16BigEndian();
                var yMin = _reader.ReadInt16BigEndian();
                var xMax = _reader.ReadInt16BigEndian();
                var yMax = _reader.ReadInt16BigEndian();

                glyphHeaders[x] = new GlyphHeader
                {
                    numberOfContours = numberOfContours,
                    xMin = xMin,
                    yMin = yMin,
                    xMax = xMax,
                    yMax = yMax
                };
            }
            return new GlyphTable
            {
                Glyphs = glyphHeaders
            };
        }
        

    }
}
