using System;
using System.Collections.Generic;
using System.Globalization;

namespace EPPlus.Fonts.OpenType.Integration.RichText
{
    /// <summary>
    /// A list of rich-text fragments with relation to eachother is a paragraph
    /// </summary>
    internal class TextParagraph
    {
        /// <summary>
        /// The Unalatered input fragments
        /// </summary>
        List<TextFragment> InputFragments;
        //The text of the entire paragraph
        //regardless of linebreaking or style runs
        string FullText;
        List<CharInfo> AllChars;
        List<int> SeparatorIndicies = new List<int>();
        List<int> ParagraphSeparatorIndicies = new List<int>();
        List<SubParagraph> SubParagraphs = new List<SubParagraph>();
        List<StyleRun> StyleRuns = new List<StyleRun>();
        int FullTextLength = 0;

        TextLineCollection WrappedLineCollection;

        internal TextParagraph(List<TextFragment> fragments, IEnumerable<string> FontDirectories)
        {
            InputFragments = fragments;
            //Extract basic info about the entire paragraph
            InitalizeAllTextAndCharInfo();
            //Split into sub paragraphs
            Segmentation();

            //TODO: Bi-directional analysis (level-runs) will need to be merged with style runs

            //Segmenting Style Runs (Itimization)
            Itemization();

            //Apply shaping (Scripting and Cluster) in simplest of terms Measure widths/heights of characters in runs
            Shaping(FontDirectories);

            ////Line-breaking
            //Wrapping(FontDirectories, dou);
        }

        /// <summary>
        ///  Extract basic info about the entire paragraph
        /// </summary>
        void InitalizeAllTextAndCharInfo()
        {
            List<int> fragmentStartIdx = new List<int>();
            int allCharIdx = 0;
            int fragmentIdx = 0;
            foreach (var fragment in InputFragments)
            {
                //var currentShaper = OpenTypeFonts.GetShaperForFont(fragment.Font);

                //var currShapedText = currentShaper.ShapeLight(fragment.Text);

                int spanIndex = 0;
                fragmentIdx++;
                foreach (var c in fragment.Text)
                {
                    var currCharInfo = new CharInfo(allCharIdx, fragmentIdx, spanIndex);
                    AllChars.Add(currCharInfo);
                    if (char.IsSeparator(c))
                    {
                        currCharInfo.IsSeparator = true;
                        SeparatorIndicies.Add(allCharIdx);
                    }

                    spanIndex++;
                    allCharIdx++;
                }
                FullText += fragment.Text;
            }
            FullTextLength = allCharIdx;
        }

        //Split paragraphs along paragraph separators
        void Segmentation()
        {
            int lastParagraphIdx = 0;
            UnicodeCategory category = UnicodeCategory.ParagraphSeparator;
            foreach (var sepIdx in SeparatorIndicies)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(FullText[sepIdx]) == category)
                {
                    ParagraphSeparatorIndicies.Add(sepIdx);
                    var section = new SubParagraph(lastParagraphIdx, sepIdx, GetFullText, GetSection);
                    SubParagraphs.Add(section);
                    lastParagraphIdx = sepIdx;
                }
            }

            var lastSection = new SubParagraph(lastParagraphIdx, FullTextLength, GetFullText, GetSection);
            SubParagraphs.Add(lastSection);
        }

        /// <summary>
        /// Seperating input into style-runs
        /// </summary>
        void Itemization()
        {
            var styleRunStartIdx = SubParagraphs[0].FullTextStart;
            var currIdx = styleRunStartIdx;
            var currFragIdx = AllChars[0].Fragment;

            for (int i = 0; i < SubParagraphs.Count; i++)
            {
                styleRunStartIdx = SubParagraphs[i].FullTextStart;
                currIdx = styleRunStartIdx;
                currFragIdx = AllChars[i].Fragment;

                for (int j = 0; j < SubParagraphs[i].Length; j++)
                {
                    currIdx += j;
                    if (AllChars[currIdx].Fragment != currFragIdx)
                    {
                        //We have moved one beyond the last char to apply the given style.
                        //Therefore -1 (unless it is on the very first idx)
                        var styleRun = new StyleRun(currFragIdx, styleRunStartIdx, Math.Max(currIdx -1, 1), GetFullText, GetSection);
                        StyleRuns.Add(styleRun);
                        //TODO: Technically this should not get its own list it should refer back here
                        SubParagraphs[i].AddStyleRun(styleRun);
                        styleRunStartIdx = currIdx;
                    }
                }
            }

            var LastRun = new StyleRun(currFragIdx, styleRunStartIdx, Math.Max(currIdx - 1, 1), GetFullText, GetSection);
            StyleRuns.Add(LastRun);
            styleRunStartIdx = currIdx;
        }

        /// <summary>
        /// Shaping (calculating widths, heights etc.)
        /// </summary>
        /// <param name="fontDirectories"></param>
        /// <param name="shapeLight">Set to false for slower more exact positioning (very rarely neccesary)</param>
        void Shaping(IEnumerable<string> fontDirectories, bool shapeLight = true)
        {
            foreach (var styleRun in StyleRuns)
            {
                var inputFrag = InputFragments[styleRun.FragmentIndex];
                var shaper = OpenTypeFonts.GetShaperForFont(inputFrag.Font, fontDirectories);
                if(shapeLight)
                {
                    var shapedGlyphs = shaper.ShapeLight(styleRun.Text);
                    double[] charWidths = new double[styleRun.Length];
                    shapedGlyphs.FillCharWidths(inputFrag.Font.Size, charWidths, styleRun.Length);
                    styleRun.SetCharWidths(charWidths);
                }
                else
                {
                    throw new NotImplementedException("Proper shaping has not been implemented here yet");
                }
            }

            var lastFragment = InputFragments[InputFragments.Count];
            var lastRun = StyleRuns[StyleRuns.Count];
            var lastShaper = OpenTypeFonts.GetShaperForFont(lastFragment.Font);
            var lastShapedGlyphs = lastShaper.ShapeLight(StyleRuns[StyleRuns.Count].Text);
            double[] lastCharWidths = new double[lastRun.Length];
            lastShapedGlyphs.FillCharWidths(lastFragment.Font.Size, lastCharWidths, lastRun.Length);
            lastRun.SetCharWidths(lastCharWidths);
        }
        /// <summary>
        /// Wrapping/line breaking
        /// </summary>
        /// <param name="fontDirectories"></param>
        /// <param name="maxWidth"></param>
        /// <returns></returns>
        internal TextLineCollection Wrap(IEnumerable<string> fontDirectories, double maxWidth)
        {
            var layoutEngine = OpenTypeFonts.GetTextLayoutEngineForFont(InputFragments[0].Font, fontDirectories);
            var wrappedLines = layoutEngine.WrapRichTextLines(InputFragments, maxWidth);
            WrappedLineCollection = new TextLineCollection(wrappedLines, InputFragments);
            return WrappedLineCollection;
        }

        string GetSection(int startIdx, int endIdx)
        {
            var subString = FullText.Substring(startIdx, endIdx - startIdx + 1);
            return subString;
        }

        string GetFullText()
        {
            return FullText;
        }

        List<CharInfo> GetCharInfoOfStyleRun(StyleRun run)
        {
            List<CharInfo> infoLst = new List<CharInfo>();
            for (int i = 0; i < run.Length; i++)
            {
                var charIdx = run.FullTextStart + i;
                infoLst.Add(AllChars[charIdx]);
            }
            return infoLst;
        }

        //Get paragraphindex

    }
}
