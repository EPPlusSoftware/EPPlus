/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           TextLayoutEngine implementation
  01/22/2025         EPPlus Software AB           Optimized with shaping cache
  01/23/2025         EPPlus Software AB           Fixed lastSpaceIndex bug in multi-fragment wrapping
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// Rich text wrapping functionality for TextLayoutEngine.
    /// </summary>
    public partial class TextLayoutEngine
    {
        /// <summary>
        /// Wraps rich text with multiple fonts without full text concatenation.
        /// Processes fragments sequentially with persistent line state.
        /// </summary>
        public List<string> WrapRichText(
            List<TextFragment> fragments,
            double maxWidthPoints)
        {
            if (fragments == null || fragments.Count == 0)
            {
                return new List<string> { string.Empty };
            }

            _lineListBuffer.Clear();

            var lineBuilder = new StringBuilder(512);
            double lineWidth = 0;
            int lastSpaceIndex = -1;

            foreach (var fragment in fragments)
            {
                if (string.IsNullOrEmpty(fragment.Text)) continue;

                ProcessFragment(fragment, maxWidthPoints, lineBuilder, ref lineWidth, ref lastSpaceIndex);
            }

            FinalizeCurrentLine(lineBuilder, lineWidth, lastSpaceIndex);

            if (_lineListBuffer.Count == 0)
            {
                _lineListBuffer.Add(string.Empty);
            }

            return new List<string>(_lineListBuffer);
        }

        private void ProcessFragment(
            TextFragment fragment,
            double maxWidthPoints,
            StringBuilder lineBuilder,
            ref double lineWidth,
            ref int lastSpaceIndex)
        {
            var shaper = GetShaperForFont(fragment.Font);
            var options = fragment.Options ?? ShapingOptions.Default;

            int len = fragment.Text.Length;

            var charWidths = GetCharWidthBuffer(len);

            var shaped = shaper.Shape(fragment.Text, options);
            double scale = fragment.Font.Size / shaper.UnitsPerEm;

            Array.Clear(charWidths, 0, len);
            FillCharWidths(shaped.Glyphs, scale, len, charWidths);

            int i = 0;
            while (i < len)
            {
                char c = fragment.Text[i];

                if (IsLineBreak(c))
                {
                    HandleLineBreak(lineBuilder, lineWidth, lastSpaceIndex);
                    SkipLineBreakChars(fragment.Text, ref i);
                    lineWidth = 0;
                    lastSpaceIndex = -1;  // Reset after line break
                    continue;
                }

                lineBuilder.Append(c);
                lineWidth += charWidths[i];

                if (c == ' ')
                {
                    lastSpaceIndex = lineBuilder.Length - 1;
                }

                if (lineWidth > maxWidthPoints)
                {
                    WrapCurrentLine(lineBuilder, lineWidth, lastSpaceIndex, maxWidthPoints);
                    lineWidth = 0;
                    lastSpaceIndex = -1;  // Reset after wrap
                }

                i++;
            }
        }

        private void FillCharWidths(ShapedGlyph[] glyphs, double scale, int textLength, double[] charWidths)
        {
            foreach (var glyph in glyphs)
            {
                int idx = glyph.ClusterIndex;
                if (idx >= 0 && idx < textLength)
                {
                    charWidths[idx] += glyph.XAdvance * scale;
                }
            }
        }

        private bool IsLineBreak(char c)
        {
            return c == '\r' || c == '\n';
        }

        private void HandleLineBreak(StringBuilder lineBuilder, double lineWidth, int lastSpaceIndex)
        {
            if (lineBuilder.Length > 0 && lineBuilder[lineBuilder.Length - 1] == ' ')
            {
                lineBuilder.Length--;
            }
            if (lineBuilder.Length > 0)
            {
                _lineListBuffer.Add(lineBuilder.ToString());
            }
            else if (lineWidth > 0)
            {
                _lineListBuffer.Add(string.Empty);
            }

            lineBuilder.Length = 0;
        }

        private void SkipLineBreakChars(string text, ref int i)
        {
            if (i < text.Length - 1 && text[i] == '\r' && text[i + 1] == '\n')
            {
                i++;
            }
            i++;
        }

        private void WrapCurrentLine(StringBuilder lineBuilder, double lineWidth, int lastSpaceIndex, double maxWidthPoints)
        {
            // Bounds check to prevent ArgumentOutOfRangeException
            if (lastSpaceIndex >= 0 && lastSpaceIndex < lineBuilder.Length)
            {
                string line = lineBuilder.ToString(0, lastSpaceIndex).TrimEnd();
                _lineListBuffer.Add(line);
                lineBuilder.Remove(0, lastSpaceIndex + 1);
            }
            else
            {
                // No valid space found - wrap entire line
                _lineListBuffer.Add(lineBuilder.ToString());
                lineBuilder.Length = 0;
            }
        }

        private void FinalizeCurrentLine(StringBuilder lineBuilder, double lineWidth, int lastSpaceIndex)
        {
            if (lineBuilder.Length > 0)
            {
                if (lineBuilder[lineBuilder.Length - 1] == ' ')
                {
                    lineBuilder.Length--;
                }
                if (lineBuilder.Length > 0)
                {
                    _lineListBuffer.Add(lineBuilder.ToString());
                }
            }
        }
    }
}