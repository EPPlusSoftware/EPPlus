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
  02/23/2026         EPPlus Software AB           Performance fix: Shape() → ShapeLight() in ProcessFragment
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Integration.RichText;
using EPPlus.Fonts.OpenType.Utilities;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// Rich text wrapping functionality for TextLayoutEngine.
    /// </summary>
    public partial class TextLayoutEngine
    {
        //public List<string> WrapRichText(
        //   List<TextFragment> fragments,
        //   double maxWidthPoints)
        //{
        //    var frags = fragments.Cast<ITextFragmentBase>().ToList();
        //    return WrapRichText(frags, maxWidthPoints);
        //}

        /// <summary>
        /// Wraps rich text with multiple fonts without full text concatenation.
        /// Processes fragments sequentially with persistent line state.
        /// </summary>
        public List<string> WrapRichText(
            IEnumerable<ITextFragmentBase> fragments,
            double maxWidthPoints)
        {
            //Potentially optimize this '.Count()' method is slow
            //Prefer the list parameter .Count but we also seemingly can't use List as input param
            if (fragments == null || fragments.Count() == 0)
            {
                return new List<string> { string.Empty };
            }

            _lineListBuffer.Clear();

            var lineBuilder = new StringBuilder(512);
            var state = new WrapStateRichText(0);
            state.WordStart = -1;
            state.LineStart = -1;

            foreach (var fragment in fragments)
            {
                if (string.IsNullOrEmpty(fragment.Text)) continue;

                ProcessFragment(fragment, maxWidthPoints, lineBuilder, state);
            }

            FinalizeCurrentLine(lineBuilder, state.CurrentLineWidth, state.WordStart, state.CurrentTextLine);
            state.EndCurrentTextLine();


            if (_lineListBuffer.Count == 0)
            {
                _lineListBuffer.Add(string.Empty);
            }

            return new List<string>(_lineListBuffer);
        }
        public List<string> WrapRichText(
               List<string> textFragments, List<MeasurementFont> fonts,
               double maxWidthPoints)
        {
            TextFragmentCollectionSimple fragmentCollection = new TextFragmentCollectionSimple(fonts, textFragments);
            return WrapRichText(fragmentCollection, maxWidthPoints);
        }


        public List<TextLineSimple> WrapRichTextLines(
            string text, MeasurementFont font,
            double maxWidthPoints)
        {
            var tCollection = new TextFragmentCollectionSimple(new List<MeasurementFont>() { font }, new List<string> { text });
            return WrapRichTextLines(tCollection, maxWidthPoints);
        }

        public TextLineCollection WrapRichTextLineCollection(
            List<ITextFragmentBase> fragments,
            double maxWidthPoints)
        {
            var innerLines = WrapRichTextLines(fragments, maxWidthPoints);
            var collection = new TextLineCollection(innerLines, fragments);
            return collection;
        }

        public TextLineCollection WrapRichTextLineCollection(List<TextFragment> fragments, double maxWidthPoints)
        {
            var frags = fragments.Cast<ITextFragmentBase>().ToList();
            return WrapRichTextLineCollection(frags, maxWidthPoints);
        }

        public List<TextLineSimple> WrapRichTextLines(List<TextFragment> fragments, double maxWidthPoints)
        {
            var frags = fragments.Cast<ITextFragmentBase>().ToList();
            return WrapRichTextLines(frags, maxWidthPoints);
        }

        public List<TextLineSimple> WrapRichTextLines(
            List<ITextFragmentBase> fragments,
            double maxWidthPoints)
        {
            if (fragments == null || fragments.Count == 0)
            {
                return new List<TextLineSimple>();
            }

            _lineListBuffer.Clear();

            var lineBuilder = new StringBuilder(512);
            var state = new WrapStateRichText(0);
            state.WordStart = -1;
            state.LineStart = -1;

            foreach (var fragment in fragments)
            {
                if (string.IsNullOrEmpty(fragment.Text)) continue;

                ProcessFragment(fragment, maxWidthPoints, lineBuilder, state);
            }

            FinalizeCurrentLine(lineBuilder, state.CurrentLineWidth, state.WordStart, state.CurrentTextLine);
            state.CurrentTextLine.Width = state.CurrentLineWidth;
            state.CurrentTextLine.Text = lineBuilder.ToString();
            state.EndCurrentTextLine();

            //Calculate ascent and descent so later application can handle line-spacing
            //This could be optimized by doing it during ProcessFragment but that is way bulkier/unclear
            foreach (var line in state.Lines)
            {
                double largestAscent = 0;
                double largestDescent = 0;
                double largestFontSize = 0;
                foreach (var lineFragment in line.InternalLineFragments)
                {
                    var frag = fragments[lineFragment.FragmentIndex];
                    if (frag == null) continue;
                    largestAscent = Math.Max(frag.AscentPoints, largestAscent);
                    largestDescent = Math.Max(frag.DescentPoints, largestDescent);
                    largestFontSize = Math.Max(largestFontSize, frag.Size);
                }
                line.LargestAscent = largestAscent;
                line.LargestDescent = largestDescent;
                line.LargestFontSize = largestFontSize;

                line.FinalizeLineFragments(fragments);
            }

            if (_lineListBuffer.Count == 0)
            {
                _lineListBuffer.Add(string.Empty);
            }

            return state.Lines;
        }

        public List<TextLineSimple> WrapRichTextRuns(
            List<StyleRun> fragments,
            double maxWidthPoints)
        {
            if (fragments == null || fragments.Count == 0)
            {
                return new List<TextLineSimple>();
            }

            _lineListBuffer.Clear();

            var lineBuilder = new StringBuilder(512);
            var state = new WrapStateRichText(0);
            state.WordStart = -1;
            state.LineStart = -1;

            foreach (var fragment in fragments)
            {
                if (string.IsNullOrEmpty(fragment.Text)) continue;

                ProcessStyleRun(fragment, maxWidthPoints, lineBuilder, state);
            }

            FinalizeCurrentLine(lineBuilder, state.CurrentLineWidth, state.WordStart, state.CurrentTextLine);
            state.CurrentTextLine.Width = state.CurrentLineWidth;
            state.CurrentTextLine.Text = lineBuilder.ToString();
            state.EndCurrentTextLine();

            if (_lineListBuffer.Count == 0)
            {
                _lineListBuffer.Add(string.Empty);
            }

            return state.Lines;
        }
        private void ProcessStyleRun(
        StyleRun run,
        double maxWidthPoints,
        StringBuilder lineBuilder,
        WrapStateRichText state)
        {
            state.CharIdxRt = 0;
            state.CharIdxWithinOriginal = run.FullTextStart;

            state.LineFrag = new LineFragment(state.CurrentFragmentIdx, lineBuilder.Length, state.CharIdxRt, state.CharIdxWithinOriginal);
            state.LineFrag.SpaceWidth = run.SpaceWidth;

            int i = 0;
            var len = run.Length;
            while (i < (len))
            {
                char c = run.Text[i];

                if (IsLineBreak(c))
                {
                    HandleLineBreak(lineBuilder, state);
                    SkipLineBreakChars(run.Text, ref i);

                    state.CurrentLineWidth = 0;
                    state.CurrentWordWidth = 0;
                    state.WordStart = -1;
                    state.LineStart = -1;
                    continue;
                }

                state.CharIdxRt = i;

                var cWidth = run.GetCharWidthByIndex(i);

                state.CurrentLineWidth += cWidth;
                state.CurrentWordWidth += cWidth;
                state.LineFrag.Width += cWidth;

                lineBuilder.Append(c);

                if (c == ' ')
                {
                    state.SetAndLogWordStartState(lineBuilder.Length - 1);
                }

                if (state.CurrentLineWidth > maxWidthPoints)
                {
                    WrapCurrentLine(lineBuilder, state, maxWidthPoints, cWidth);
                }
                i++;
                state.CharIdxWithinOriginal++;
                state.CharIdxRt = i;
            }

            if (state.LineFrag.Width > 0)
            {
                state.CurrentTextLine.InternalLineFragments.Add(state.LineFrag);
            }

            state.CurrentFragmentIdx++;
        }

        private void ProcessFragment(
            ITextFragmentBase fragment,
            double maxWidthPoints,
            StringBuilder lineBuilder,
            WrapStateRichText state)
        {
            state.CharIdxRt = 0;
            var shaper = GetShaperForFont((IFontFormatBase)fragment.RichTextOptions);
            var options = fragment.Options ?? ShapingOptions.Default;
            int len = fragment.Text.Length;
            var charWidths = GetCharWidthBuffer(len);
            Array.Clear(charWidths, 0, len);

            // ShapeLight applies only kerning (sufficient for line-breaking).
            // Full Shape() runs SingleAdjustment + Kerning + MarkToBase which
            // is ~250x slower and irrelevant for wrapping decisions.
            var shaped = shaper.ShapeLight(fragment.Text, options);
            shaped.FillCharWidths(fragment.Size, charWidths, len);

            //Store for after everything is done
            fragment.AscentPoints = shaper.GetAscentInPoints(fragment.Size);
            fragment.DescentPoints = shaper.GetDescentInPoints(fragment.Size);

            var spaceWidth = shaper.Shape(" ", options).GetWidthInPoints(fragment.Size);
            state.LineFrag = new LineFragment(state.CurrentFragmentIdx, lineBuilder.Length, state.CharIdxRt, state.CharIdxWithinOriginal);
            state.LineFrag.SpaceWidth = spaceWidth;

            int i = 0;
            while (i < len)
            {
                char c = fragment.Text[i];

                if (IsLineBreak(c))
                {
                    HandleLineBreak(lineBuilder, state);
                    SkipLineBreakChars(fragment.Text, ref i);

                    state.CurrentLineWidth = 0;
                    state.CurrentWordWidth = 0;
                    state.WordStart = -1;
                    state.LineStart = -1;
                    continue;
                }

                state.CharIdxRt = i;

                state.CurrentLineWidth += charWidths[i];
                state.CurrentWordWidth += charWidths[i];
                state.LineFrag.Width += charWidths[i];

                lineBuilder.Append(c);

                if (c == ' ')
                {
                    state.SetAndLogWordStartState(lineBuilder.Length - 1);
                }

                if (state.CurrentLineWidth > maxWidthPoints)
                {
                    WrapCurrentLine(lineBuilder, state, maxWidthPoints, charWidths[i]);
                }
                i++;
                state.CharIdxWithinOriginal++;
            }

            if (state.LineFrag.Width > 0)
            {
                state.CurrentTextLine.InternalLineFragments.Add(state.LineFrag);
            }

            state.CurrentFragmentIdx++;
        }

        /// <summary>
        /// Fills character widths from lightweight GlyphWidth structs (8 bytes each).
        /// Used by the wrapping pipeline for optimal performance.
        /// </summary>
        private void FillCharWidths(GlyphWidth[] glyphs, double scale, int textLength, double[] charWidths)
        {
            for (int i = 0; i < glyphs.Length; i++)
            {
                int idx = glyphs[i].ClusterIndex;
                if (idx >= 0 && idx < textLength)
                {
                    charWidths[idx] += glyphs[i].XAdvance * scale;
                }
            }
        }

        private bool IsLineBreak(char c)
        {
            return c == '\r' || c == '\n';
        }

        private void HandleLineBreak(StringBuilder lineBuilder, WrapStateRichText state)
        {
            if (lineBuilder.Length > 0 && lineBuilder[lineBuilder.Length - 1] == ' ')
            {
                lineBuilder.Length--;
            }
            if (lineBuilder.Length > 0)
            {
                _lineListBuffer.Add(lineBuilder.ToString());
                state.CurrentTextLine.Text = lineBuilder.ToString();
            }
            else if (state.CurrentLineWidth > 0)
            {
                _lineListBuffer.Add(string.Empty);
                state.CurrentTextLine.Text = string.Empty;
            }

            state.CurrentTextLine.Width = state.CurrentLineWidth;
            state.EndCurrentTextLineAndIntializeNext(0);

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
        private void WrapCurrentLine(StringBuilder lineBuilder, WrapStateRichText state, double maxWidthPoints, double advanceWidth)
        {
            int fragIdxAtBreak = state.CurrentFragmentIdx;

            int adjustmentForLineBuilderLength = 0;

            // Bounds check to prevent ArgumentOutOfRangeException
            if (state.WordStart >= 0 && state.WordStart < lineBuilder.Length)
            {
                var lineStringWithTrail = lineBuilder.ToString(0, state.WordStart + 1);//+1 was just added and should be here but everything else is sorta based on it being gone...
                if (lineStringWithTrail[lineStringWithTrail.Length - 1] == ' ')
                {
                    state.CurrentTextLine.WasWrappedOnSpace = true;
                }
                string line = lineStringWithTrail.TrimEnd();
                _lineListBuffer.Add(line);
                lineBuilder.Remove(0, state.WordStart + 1);

                //handle line data
                state.CurrentTextLine.Width = state.CurrentLineWidth - state.CurrentWordWidth;
                state.CurrentTextLine.Text = line;

                fragIdxAtBreak = state.GetFragIdxAtWordStart();
                //Because of word-wrap we may have richTextFragments on the current line that is no longer part of it after wrap.
                state.AdjustLineFragmentsForNextLine();

                state.CurrentLineWidth = state.CurrentWordWidth;
                state.LineStart = state.WordStart;
            }
            else
            {
                var lastChar = lineBuilder[lineBuilder.Length - 1];
                var line = lineBuilder.ToString(0, lineBuilder.Length - 1);
                // No valid space found - wrap entire line
                _lineListBuffer.Add(line);

                //handle line data
                state.CurrentTextLine.Width = state.CurrentLineWidth - advanceWidth;
                state.CurrentTextLine.Text = line;
                //state.CurrentTextLine.

                //Add the char that went over max to the next line
                state.CurrentLineWidth = 0;
                lineBuilder.Length = 0;

                //Append the char that goes over max unless it is a space
                if (lastChar != ' ')
                {
                    lineBuilder.Append(lastChar);
                    //Since we appended we should remove it from line builder length when end current and initialize next happens
                    adjustmentForLineBuilderLength = 1;

                    //The char that made us move past maxWidth
                    //must be added to the new line
                    state.CurrentWordWidth = advanceWidth;
                    state.CurrentLineWidth = advanceWidth;
                }
                else
                {
                    state.CurrentTextLine.WasWrappedOnSpace = true;
                }
            }

            state.EndCurrentTextLineAndIntializeNext(lineBuilder.Length - adjustmentForLineBuilderLength);
            state.CurrentWordWidth = state.CurrentLineWidth;

            state.WordStart = -1;
            state.LineStart = -1;
        }

        private void FinalizeCurrentLine(StringBuilder lineBuilder, double lineWidth, int lastSpaceIndex, TextLineSimple currentLine)
        {
            if (lineBuilder.Length > 0)
            {
                if (lineBuilder[lineBuilder.Length - 1] == ' ')
                {
                    currentLine.WasWrappedOnSpace = true;
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