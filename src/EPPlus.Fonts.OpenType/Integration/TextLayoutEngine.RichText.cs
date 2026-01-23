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
        /// Wraps rich text with multiple fonts.
        /// Returns list of wrapped lines as strings (font information is implicit from original fragments).
        /// </summary>
        /// <param name="fragments">Text fragments with their fonts</param>
        /// <param name="maxWidthPoints">Maximum line width in points</param>
        /// <returns>List of wrapped lines</returns>
        public List<string> WrapRichText(
            List<TextFragment> fragments,
            double maxWidthPoints)
        {
            if (fragments == null || fragments.Count == 0)
            {
                return new List<string> { string.Empty };
            }

            // Build full text and track fragment positions
            var fullTextBuilder = new System.Text.StringBuilder();
            var fragmentPositions = new List<FragmentPosition>();

            int currentPosition = 0;
            foreach (var fragment in fragments)
            {
                if (string.IsNullOrEmpty(fragment.Text))
                    continue;

                fragmentPositions.Add(new FragmentPosition
                {
                    StartIndex = currentPosition,
                    EndIndex = currentPosition + fragment.Text.Length,
                    Font = fragment.Font,
                    Options = fragment.Options ?? ShapingOptions.Default
                });

                fullTextBuilder.Append(fragment.Text);
                currentPosition += fragment.Text.Length;
            }

            string fullText = fullTextBuilder.ToString();

            if (string.IsNullOrEmpty(fullText))
            {
                return new List<string> { string.Empty };
            }

            // Split by line breaks and track paragraph positions in original text
            var paragraphs = fullText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var allLines = new List<string>();

            int paragraphStartPos = 0;
            foreach (var paragraph in paragraphs)
            {
                if (string.IsNullOrEmpty(paragraph))
                {
                    allLines.Add(string.Empty);
                    // Account for the line break character(s) that were removed
                    paragraphStartPos += GetLineBreakLength(fullText, paragraphStartPos);
                    continue;
                }

                int paragraphEndPos = paragraphStartPos + paragraph.Length;

                // Extract fragments that overlap with this paragraph
                var paragraphFragments = GetFragmentsForRange(
                    fragmentPositions,
                    paragraphStartPos,
                    paragraphEndPos);

                var wrappedLines = WrapRichParagraph(paragraph, paragraphFragments, maxWidthPoints);
                allLines.AddRange(wrappedLines);

                // Move to next paragraph (add line break length)
                paragraphStartPos = paragraphEndPos + GetLineBreakLength(fullText, paragraphEndPos);
            }

            return allLines;
        }

        /// <summary>
        /// Gets the length of line break at the specified position (1 for \n or \r, 2 for \r\n, 0 if none).
        /// </summary>
        private int GetLineBreakLength(string text, int pos)
        {
            if (pos >= text.Length)
                return 0;

            if (pos < text.Length - 1 && text[pos] == '\r' && text[pos + 1] == '\n')
                return 2;

            if (text[pos] == '\r' || text[pos] == '\n')
                return 1;

            return 0;
        }

        /// <summary>
        /// Extracts fragments that overlap with the specified text range and adjusts their positions
        /// to be relative to the range start.
        /// </summary>
        private List<FragmentPosition> GetFragmentsForRange(
            List<FragmentPosition> allFragments,
            int rangeStart,
            int rangeEnd)
        {
            var result = new List<FragmentPosition>();

            foreach (var fragment in allFragments)
            {
                // Check if fragment overlaps with range
                if (fragment.EndIndex <= rangeStart || fragment.StartIndex >= rangeEnd)
                {
                    continue; // No overlap
                }

                // Calculate overlap
                int overlapStart = Math.Max(fragment.StartIndex, rangeStart);
                int overlapEnd = Math.Min(fragment.EndIndex, rangeEnd);

                // Create new fragment with positions adjusted to be relative to range start
                result.Add(new FragmentPosition
                {
                    StartIndex = overlapStart - rangeStart,
                    EndIndex = overlapEnd - rangeStart,
                    Font = fragment.Font,
                    Options = fragment.Options
                });
            }

            return result;
        }

        /// <summary>
        /// Wraps a single rich-text paragraph (no line breaks).
        /// OPTIMIZED: Reuses _charWidthBuffer and _lineListBuffer.
        /// Uses StringBuilder for line building to minimize string allocations.
        /// </summary>
        private List<string> WrapRichParagraph(
            string text,
            List<FragmentPosition> fragmentPositions,
            double maxWidthPoints)
        {
            _lineListBuffer.Clear();

            if (string.IsNullOrEmpty(text))
            {
                _lineListBuffer.Add(string.Empty);
                return new List<string>(_lineListBuffer);
            }

            // Reuse char width buffer
            int required = text.Length;
            if (_charWidthBuffer.Length < required)
            {
                int newSize = Math.Max(required, _charWidthBuffer.Length * 2);
                Array.Resize(ref _charWidthBuffer, newSize);
            }
            Array.Clear(_charWidthBuffer, 0, required);

            // Fill charWidths from all fragments
            foreach (var fragment in fragmentPositions)
            {
                int start = fragment.StartIndex;
                int length = fragment.EndIndex - start;
                if (length <= 0) continue;

                string fragText = text.Substring(start, length);
                var shaper = GetShaperForFont(fragment.Font);
                var shaped = shaper.Shape(fragText, fragment.Options ?? ShapingOptions.Default);
                double scaleFactor = fragment.Font.Size / shaper.UnitsPerEm;

                foreach (var glyph in shaped.Glyphs)
                {
                    int idx = glyph.ClusterIndex;
                    if (idx >= 0 && idx < length)
                    {
                        _charWidthBuffer[start + idx] += glyph.XAdvance * scaleFactor;
                    }
                }
            }

            double spaceWidth = fragmentPositions.Count > 0
                ? MeasureTextWithFont(" ", fragmentPositions[0].Font, fragmentPositions[0].Options)
                : 0;

            int lineStart = 0;
            int wordStart = 0;
            double currentLineWidth = 0;
            double currentWordWidth = 0;

            var currentLineBuilder = new StringBuilder(text.Length / 4 + 20);  // Estimate for line length

            for (int i = 0; i <= text.Length; i++)
            {
                bool isSpace = (i < text.Length && text[i] == ' ');
                bool isEnd = (i == text.Length);

                if ((isSpace || isEnd) && wordStart < i)
                {
                    double totalWidth = currentLineWidth + currentWordWidth;
                    if (lineStart < wordStart)
                    {
                        var frag = GetFragmentAtPosition(i, fragmentPositions);
                        totalWidth += MeasureTextWithFont(" ", frag.Font, frag.Options);
                    }

                    if (totalWidth <= maxWidthPoints || lineStart == wordStart)
                    {
                        // Word fits - append to builder
                        if (currentLineBuilder.Length > 0)
                        {
                            currentLineBuilder.Append(' ');
                        }
                        currentLineBuilder.Append(text, wordStart, i - wordStart);
                        currentLineWidth = totalWidth;
                    }
                    else
                    {
                        // Wrap - add line and start new
                        if (currentLineBuilder.Length > 0 && currentLineBuilder[currentLineBuilder.Length - 1] == ' ')
                        {
                            currentLineBuilder.Length--;
                        }
                        if (currentLineBuilder.Length > 0)
                        {
                            _lineListBuffer.Add(currentLineBuilder.ToString());
                        }
                        currentLineBuilder.Length = 0;

                        lineStart = wordStart;
                        currentLineWidth = currentWordWidth;

                        currentLineBuilder.Append(text, wordStart, i - wordStart);
                    }

                    wordStart = i + 1;
                    currentWordWidth = 0;
                }
                else if (isSpace)
                {
                    wordStart = i + 1;
                }
                else
                {
                    currentWordWidth += _charWidthBuffer[i];

                    if (currentWordWidth > maxWidthPoints && lineStart < wordStart && currentLineWidth > 0)
                    {
                        if (currentLineBuilder.Length > 0 && currentLineBuilder[currentLineBuilder.Length - 1] == ' ')
                        {
                            currentLineBuilder.Length--;
                        }
                        if (currentLineBuilder.Length > 0)
                        {
                            _lineListBuffer.Add(currentLineBuilder.ToString());
                        }
                        currentLineBuilder.Length = 0;

                        lineStart = wordStart;
                        currentLineWidth = 0;
                    }
                }
            }

            // Final line
            if (lineStart < text.Length)
            {
                if (currentLineBuilder.Length > 0 && currentLineBuilder[currentLineBuilder.Length - 1] == ' ')
                {
                    currentLineBuilder.Length--;
                }
                if (currentLineBuilder.Length > 0)
                {
                    _lineListBuffer.Add(currentLineBuilder.ToString());
                }
            }

            if (_lineListBuffer.Count == 0)
            {
                _lineListBuffer.Add(string.Empty);
            }

            return new List<string>(_lineListBuffer);
        }

        /// <summary>
        /// Finds which fragment a character position belongs to.
        /// </summary>
        private FragmentPosition GetFragmentAtPosition(int position, List<FragmentPosition> fragments)
        {
            foreach (var fragment in fragments)
            {
                if (position >= fragment.StartIndex && position < fragment.EndIndex)
                {
                    return fragment;
                }
            }

            // Fallback to last fragment
            return fragments[fragments.Count - 1];
        }

    }
}