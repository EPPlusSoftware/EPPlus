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


using EPPlus.Fonts.OpenType.Tables.Glyph.Serialization;

namespace EPPlus.Fonts.OpenType.Tables.Glyph
{
    internal class GlyfTableLoader : TableLoader<GlyfTable>
    {
        private readonly uint[] _glyphOffsets;
        private readonly BoundingRectangle _emptyGlyph;

        public GlyfTableLoader(TableLoaderSettings settings) : base(settings, TableNames.Glyf)
        {
            _glyphOffsets = TableLoaders.GetLocaTableLoader(settings).Load().Offsets;
            _emptyGlyph = TableLoaders.GetHeadTableLoader(settings).Load().GetDefaultBounds();
            _settings = settings;
        }


        private readonly TableLoaderSettings _settings;

        protected override GlyfTable LoadInternal()
        {
            var glyphs = new Glyph[_glyphOffsets.Length];
            _reader.SetContext("glyf");
            for (int i = 0; i < _glyphOffsets.Length - 1; i++)
            {
                var start = _glyphOffsets[i];
                var end = _glyphOffsets[i + 1];

                if (start == end)
                {
                    continue;
                }

                _reader.BaseStream.Position = _offset + start;
                glyphs[i] = GlyphDeserializer.Deserialize(_reader);
            }
            _reader.SetContext(string.Empty);
            return new GlyfTable(_settings) { Glyphs = glyphs };
        }
    }
}
