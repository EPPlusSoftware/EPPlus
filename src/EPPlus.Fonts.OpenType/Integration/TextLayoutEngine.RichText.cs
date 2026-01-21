using OfficeOpenXml.Interfaces.Drawing.Text;
using EPPlus.Fonts.OpenType.TextShaping;
using System;
using System.Collections.Generic;

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

            // Build full paragraph text and track fragment positions
            var paragraphBuilder = new System.Text.StringBuilder();
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

                paragraphBuilder.Append(fragment.Text);
                currentPosition += fragment.Text.Length;
            }

            string fullText = paragraphBuilder.ToString();

            if (string.IsNullOrEmpty(fullText))
            {
                return new List<string> { string.Empty };
            }

            // Handle existing line breaks
            var paragraphs = fullText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            var allLines = new List<string>();

            foreach (var paragraph in paragraphs)
            {
                if (string.IsNullOrEmpty(paragraph))
                {
                    allLines.Add(string.Empty);
                    continue;
                }

                var wrappedLines = WrapRichParagraph(paragraph, fragmentPositions, maxWidthPoints);
                allLines.AddRange(wrappedLines);
            }

            return allLines;
        }

        /// <summary>
        /// Wraps a single rich-text paragraph (no line breaks).
        /// </summary>
        private List<string> WrapRichParagraph(
            string text,
            List<FragmentPosition> fragmentPositions,
            double maxWidthPoints)
        {
            var lines = new List<string>();

            var currentLine = new System.Text.StringBuilder();
            var currentWord = new System.Text.StringBuilder();

            double currentLineWidth = 0;
            double currentWordWidth = 0;
            int wordStartIndex = 0;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                if (c == ' ')
                {
                    // Try to add word + space to current line
                    string wordText = currentWord.ToString();

                    // Measure space in the font at this position
                    var fragment = GetFragmentAtPosition(i, fragmentPositions);
                    double spaceWidth = MeasureTextWithFont(" ", fragment.Font, fragment.Options);

                    double totalWidth = currentLineWidth + currentWordWidth + spaceWidth;

                    if (totalWidth <= maxWidthPoints || currentLine.Length == 0)
                    {
                        // Word fits on current line
                        if (currentLine.Length > 0)
                        {
                            currentLine.Append(' ');
                            currentLineWidth += spaceWidth;
                        }
                        currentLine.Append(wordText);
                        currentLineWidth += currentWordWidth;

                        // Reset word
                        currentWord.Length = 0;
                        currentWordWidth = 0;
                    }
                    else
                    {
                        // Word doesn't fit - wrap to new line
                        lines.Add(currentLine.ToString());

                        currentLine.Length = 0;
                        currentLine.Append(wordText);
                        currentLineWidth = currentWordWidth;

                        currentWord.Length = 0;
                        currentWordWidth = 0;
                    }

                    wordStartIndex = i + 1; // Next word starts after space
                }
                else
                {
                    // Add character to current word
                    currentWord.Append(c);

                    // Measure word so far (may span multiple fonts)
                    string wordSoFar = currentWord.ToString();
                    currentWordWidth = MeasureWordAcrossFragments(
                        wordSoFar,
                        wordStartIndex,
                        fragmentPositions);

                    // Check if word itself is too long for a line
                    if (currentLineWidth + currentWordWidth > maxWidthPoints && currentLine.Length > 0)
                    {
                        // Wrap current line and start new line with this word
                        lines.Add(currentLine.ToString());
                        currentLine.Length = 0;
                        currentLineWidth = 0;
                    }
                }
            }

            // Add remaining word and line
            if (currentWord.Length > 0)
            {
                string wordText = currentWord.ToString();

                if (currentLine.Length > 0 && currentLineWidth + currentWordWidth > maxWidthPoints)
                {
                    // Word doesn't fit - wrap to new line
                    lines.Add(currentLine.ToString());
                    currentLine.Length = 0;
                    currentLine.Append(wordText);
                }
                else
                {
                    // Word fits
                    if (currentLine.Length > 0)
                    {
                        currentLine.Append(' ');
                    }
                    currentLine.Append(wordText);
                }
            }

            if (currentLine.Length > 0)
            {
                lines.Add(currentLine.ToString());
            }

            // Ensure at least one line
            if (lines.Count == 0)
            {
                lines.Add(string.Empty);
            }

            return lines;
        }

        /// <summary>
        /// Measures a word that may span multiple fragments with different fonts.
        /// </summary>
        private double MeasureWordAcrossFragments(
            string word,
            int wordStartIndex,
            List<FragmentPosition> fragments)
        {
            if (string.IsNullOrEmpty(word))
            {
                return 0;
            }

            double totalWidth = 0;
            int wordEndIndex = wordStartIndex + word.Length;

            // Iterate through fragments to find which ones overlap with this word
            foreach (var fragment in fragments)
            {
                // Check if this fragment overlaps with the word
                if (fragment.EndIndex <= wordStartIndex || fragment.StartIndex >= wordEndIndex)
                {
                    continue; // No overlap
                }

                // Calculate overlap
                int overlapStart = Math.Max(fragment.StartIndex, wordStartIndex);
                int overlapEnd = Math.Min(fragment.EndIndex, wordEndIndex);

                // Extract the portion of the word in this fragment
                int localStart = overlapStart - wordStartIndex;
                int localEnd = overlapEnd - wordStartIndex;
                string section = word.Substring(localStart, localEnd - localStart);

                // Measure this section with the fragment's font
                double sectionWidth = MeasureTextWithFont(section, fragment.Font, fragment.Options);
                totalWidth += sectionWidth;
            }

            return totalWidth;
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