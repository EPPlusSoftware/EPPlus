using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EPPlus.Fonts.OpenType.Utilities;
using OfficeOpenXml.Interfaces.Fonts;


namespace EPPlus.Fonts.OpenType.Integration
{
    public partial class TextLayoutEngine
    {
        private List<string> CreateEmptyResult()
        {
            _lineListBuffer.Clear();
            _lineListBuffer.Add(string.Empty);
            return new List<string>(_lineListBuffer);
        }

        private void PrepareLineBuilder(int textLength)
        {
            _lineBuilder.Clear();
            _lineBuilder.EnsureCapacity(textLength / 4 + 20);
        }

        private double[] CalculateCharacterWidths(string text, float fontSize, ShapingOptions options)
        {
            var glyphs = _shaper.ShapeLight(text, options);
            var charWidths = GetCharWidthBuffer(text.Length);
            Array.Clear(charWidths, 0, text.Length);

            double scaleFactor = fontSize / _shaper.UnitsPerEm;

            foreach (var glyph in glyphs)
            {
                int charIndex = glyph.ClusterIndex;
                if (charIndex >= 0 && charIndex < text.Length)
                {
                    charWidths[charIndex] += glyph.XAdvance * scaleFactor;
                }
            }

            return charWidths;
        }



        private CharacterType GetCharacterType(string text, int position)
        {
            if (position >= text.Length)
                return CharacterType.EndOfText;

            if (text[position] == ' ')
                return CharacterType.Space;

            return CharacterType.Regular;
        }
    }
}
