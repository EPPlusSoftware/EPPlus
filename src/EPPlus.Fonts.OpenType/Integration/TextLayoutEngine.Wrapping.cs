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
using EPPlus.Fonts.OpenType.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    public partial class TextLayoutEngine
    {
        /// <summary>
        /// Processes a complete word (reached space or end of text).
        /// Decides whether to add it to current line or start a new line.
        /// </summary>
        private void ProcessCompleteWord(
            string text,
            WrapStateText state,
            int currentPos,
            double maxWidth)
        {
            double totalWidth = state.CurrentLineWidth + state.CurrentWordWidth;

            totalWidth += state.SpaceWidth;

            if (totalWidth <= maxWidth || state.LineStart == state.WordStart)
            {
                // Word fits on current line
                _lineBuilder.AppendSpaceIfNotEmpty();
                _lineBuilder.AppendSubstring(text, state.WordStart, currentPos - state.WordStart);
                state.CurrentLineWidth = totalWidth;
            }
            else
            {
                // Word doesn't fit - start new line
                _lineBuilder.FlushToList(_lineListBuffer);
                _lineListBuffer[_lineListBuffer.Count-1] += " ";

                state.LineStart = state.WordStart;
                state.CurrentLineWidth = state.CurrentWordWidth;

                _lineBuilder.AppendSubstring(text, state.WordStart, currentPos - state.WordStart);
            }

            state.WordStart = currentPos + 1;
            state.CurrentWordWidth = 0;
        }

        private void ProcessNonEndingSpace(string text,
            WrapStateText state,
            int currentPos,
            double maxWidth
            )
        {
            var totalWidth = state.CurrentLineWidth + state.SpaceWidth;

            if (totalWidth <= maxWidth)
            {
                // Space fits on current line
                _lineBuilder.AppendSpaceIfNotEmpty();
                _lineBuilder.AppendSubstring(text, state.WordStart, currentPos - state.WordStart);
                state.CurrentLineWidth = totalWidth;
            }
            else
            {
                // Word doesn't fit - start new line
                _lineBuilder.FlushToList(_lineListBuffer);

                state.LineStart = state.WordStart;
                state.CurrentLineWidth = state.CurrentWordWidth;

                _lineBuilder.AppendSubstring(text, state.WordStart, currentPos - state.WordStart);
            }

            state.WordStart = currentPos + 1;
            //Word width likely always 0 here before and after
            //state.CurrentWordWidth = 0;
        }

        private void ProcessCharacterInWord(
             string text,
             double[] charWidths,
             WrapStateText state,
             int currentPos,
             double maxWidth)
        {
            state.CurrentWordWidth += charWidths[currentPos];

            // CASE 1: Line has content and word grows too large
            if ((state.CurrentWordWidth + state.CurrentLineWidth) > maxWidth &&
                state.LineStart < state.WordStart &&
                state.CurrentLineWidth > 0)
            {
                _lineBuilder.FlushToList(_lineListBuffer);
                state.LineStart = state.WordStart;
                state.CurrentLineWidth = 0;
            }

            // CASE 2: Word is alone on line and too long - break it
            if (state.LineStart == state.WordStart && state.CurrentWordWidth > maxWidth)
            {
                BreakLongWord(text, charWidths, state, currentPos, maxWidth);
            }
        }

        /// <summary>
        /// Breaks a word that is too long to fit on a single line.
        /// Uses backward removal strategy: removes characters from the end until the word fits.
        /// </summary>
        private void BreakLongWord(
            string text,
            double[] charWidths,
            WrapStateText state,
            int currentPos,
            double maxWidth)
        {
            int breakPoint = currentPos + 1;
            double currentWidth = state.CurrentWordWidth;

            // Remove characters from the end until it fits
            while (breakPoint > state.WordStart + 1 && currentWidth > maxWidth)
            {
                breakPoint--;
                currentWidth -= charWidths[breakPoint];
            }

            // Safety: at least 1 character must fit
            if (breakPoint <= state.WordStart)
            {
                breakPoint = state.WordStart + 1;
            }

            // Add what fits on current line
            _lineBuilder.AppendSubstring(text, state.WordStart, breakPoint - state.WordStart);
            _lineListBuffer.Add(_lineBuilder.ToString());
            _lineBuilder.Clear();

            // Calculate width of remaining part
            state.CurrentWordWidth = 0;
            for (int j = breakPoint; j <= currentPos; j++)
            {
                if (j < text.Length)
                {
                    state.CurrentWordWidth += charWidths[j];
                }
            }

            // Update state
            state.WordStart = breakPoint;
            state.LineStart = breakPoint;
            state.CurrentLineWidth = 0;
        }

        private List<string> FinalizeWrapping()
        {
            _lineBuilder.FlushToList(_lineListBuffer);

            if (_lineListBuffer.Count == 0)
            {
                _lineListBuffer.Add(string.Empty);
            }

            return new List<string>(_lineListBuffer);
        }
    }
}
