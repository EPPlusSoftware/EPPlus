using OfficeOpenXml.Interfaces.RichText;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Integration.RichText
{
    /// <summary>
    /// A list of rich-text fragments with relation to eachother
    /// Essentially a paragraph handler
    /// </summary>
    public class LayoutSystem
    {
        /// <summary>
        /// The Unaltered input fragments
        /// </summary>
        public List<ITextFragmentBase> InputFragments { get; private set; }

        //The text of the entire paragraph
        //regardless of linebreaking or style runs
        string FullText;
        private List<CharInfo> AllChars { get; set; }= new List<CharInfo>();
        List<int> SeparatorIndicies = new List<int>();
        List<int> ParagraphSeparatorIndicies = new List<int>();
        List<Paragraph> SubParagraphs = new List<Paragraph>();
        private List<StyleRun> StyleRuns { get; set; } = new List<StyleRun>();
        int FullTextLength = 0;
        int FullTextLastIdx = 0;


        TextLineCollection WrappedLineCollection;

        public LayoutSystem(List<TextFragment> preFragments): this(preFragments.Cast<ITextFragmentBase>().ToList())
        {
        }

        public LayoutSystem(IEnumerable<IRichTextFormatEssential> preFragments)
        {
            InputFragments = new List<ITextFragmentBase>();

            foreach(var preFrag in preFragments)
            {
                var frag = new TextFragmentBase(preFrag);
                InputFragments.Add(frag);
            }
            InitializeLayout();
        }

        public LayoutSystem(IEnumerable<ITextFragmentBase> fragments)
        {
            InputFragments = fragments.ToList();
            InitializeLayout();
        }

        private void InitializeLayout()
        {
            if (InputFragments != null && InputFragments.Count > 0)
            {
                //Extract basic info about the entire paragraph
                InitalizeAllTextAndCharInfo();
                //Split into sub paragraphs
                Segmentation();

                //TODO: Bi-directional analysis (level-runs) these will need to be merged with style runs

                //Segmenting Style Runs (Itimization)
                Itemization();

                //Apply shaping (Scripting and Cluster) in simplest of terms Measure widths/heights of characters in runs
                Shaping();
            }
            //Line-breaking (Is called by user instead)
            //Wrap();
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
                int spanIndex = 0;
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
                fragmentIdx++;
            }
            FullTextLength = allCharIdx;
            FullTextLastIdx = FullTextLength - 1;
            if(FullTextLength != 0 && FullTextLastIdx == 0)
            {
                FullTextLastIdx = 1;
            }
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
                    var section = new Paragraph(lastParagraphIdx, sepIdx-1, GetFullText, GetSection);
                    SubParagraphs.Add(section);
                    lastParagraphIdx = sepIdx;
                }
            }

            var lastSection = new Paragraph(lastParagraphIdx, FullTextLength-1, GetFullText, GetSection);
            SubParagraphs.Add(lastSection);
        }

        /// <summary>
        /// Seperating input into style-runs
        /// </summary>
        void Itemization()
        {
            var subParagraphStartIdx = SubParagraphs[0].FullTextStart;
            var currIdx = subParagraphStartIdx;
            var currFragIdx = 0;
            if (AllChars != null && AllChars.Count != 0)
            {
                currFragIdx = AllChars[0].Fragment;

                var lastRunIdx = 0;

                for (int i = 0; i < SubParagraphs.Count; i++)
                {
                    subParagraphStartIdx = SubParagraphs[i].FullTextStart;
                    currIdx = subParagraphStartIdx;

                    currFragIdx = AllChars[i].Fragment;

                    for (int j = 0; j < SubParagraphs[i].Length; j++)
                    {
                        currIdx = subParagraphStartIdx + j;
                        if (AllChars[currIdx].Fragment != currFragIdx)
                        {
                            //We have moved one beyond the last char to apply the given style.
                            //Therefore -1 (unless it is on the very first idx)
                            var styleRun = new StyleRun(currFragIdx, lastRunIdx, currIdx -1, GetFullText, GetSection);
                            StyleRuns.Add(styleRun);
                            //TODO: Technically this should not get its own list it should refer back here
                            SubParagraphs[i].AddStyleRun(styleRun);
                            currFragIdx = AllChars[currIdx].Fragment;
                            lastRunIdx = currIdx;
                        }
                    }

                }

                var LastRun = new StyleRun(currFragIdx, lastRunIdx, currIdx, GetFullText, GetSection);
                StyleRuns.Add(LastRun);
                SubParagraphs[SubParagraphs.Count - 1].AddStyleRun(LastRun);
            }
        }

        /// <summary>
        /// Shaping (calculating widths, heights etc.)
        /// </summary>
        /// <param name="fontDirectories"></param>
        /// <param name="shapeLight">Set to false for slower more exact positioning (very rarely neccesary)</param>
        void Shaping(bool shapeLight = true)
        {
            if (InputFragments != null && InputFragments.Count > 0)
            {
                foreach (var styleRun in StyleRuns)
                {
                    var inputFrag = InputFragments[styleRun.FragmentIndex];
                    var shaper = OpenTypeFonts.GetTextShaper(inputFrag.RichTextOptions.Family, inputFrag.RichTextOptions.SubFamily);

                    if (shapeLight)
                    {
                        var shapedGlyphs = shaper.ShapeLight(styleRun.Text);

                        double[] charWidths = new double[styleRun.Length + 1];
                        shapedGlyphs.FillCharWidths((float)inputFrag.RichTextOptions.Size, charWidths, styleRun.Length + 1);
                        var spaceWidth = shaper.Shape(" ").GetWidthInPoints((float)inputFrag.RichTextOptions.Size);

                        inputFrag.AscentPoints = shaper.GetAscentInPoints(inputFrag.RichTextOptions.Size);
                        inputFrag.DescentPoints = shaper.GetDescentInPoints(inputFrag.RichTextOptions.Size);

                        int charIdx = styleRun.FullTextStart;
                        foreach (var width in charWidths)
                        {
                            if (charIdx < AllChars.Count)
                            {
                                AllChars[charIdx].Width = width;
                                charIdx++;
                            }
                        }

                        styleRun.SetCharWidths(charWidths, spaceWidth);
                    }
                    else
                    {
                        throw new NotImplementedException("Proper shaping has not been implemented here yet");
                    }
                }

                var lastFragment = InputFragments[InputFragments.Count - 1];
                var lastRun = StyleRuns[StyleRuns.Count - 1];
                var lastShaper = OpenTypeFonts.GetTextShaper(lastFragment.RichTextOptions.Family, lastFragment.RichTextOptions.SubFamily);
                var lastShapedGlyphs = lastShaper.ShapeLight(lastRun.Text);
                double[] lastCharWidths = new double[lastRun.Length + 1];
                lastShapedGlyphs.FillCharWidths((float)lastFragment.RichTextOptions.Size, lastCharWidths, lastRun.Length + 1);
                var LastspaceWidth = lastShaper.Shape(" ").GetWidthInPoints((float)lastFragment.RichTextOptions.Size);

                lastRun.SetCharWidths(lastCharWidths, LastspaceWidth);
            }
        }

        /// <summary>
        /// Wrapping/line breaking
        /// </summary>
        /// <param name="fontDirectories"></param>
        /// <param name="maxWidth"></param>
        /// <returns></returns>
        public TextLineCollection Wrap(double maxWidth)
        {
            if (InputFragments == null || InputFragments.Count <= 0)
            {
                return new TextLineCollection();
            }
            var inputRt = InputFragments[0];
            var shaper = OpenTypeFonts.GetTextShaper(inputRt.RichTextOptions.Family, inputRt.RichTextOptions.SubFamily);
            var layoutEngine = new TextLayoutEngine(shaper);
            var wrappedLines = layoutEngine.WrapRichTextRuns(StyleRuns, maxWidth);

            if(wrappedLines.Count > 1)
            {
                //This could be done during wrapping technically for optimization
                for (int i = 1; i < wrappedLines.Count; i++)
                {
                    var startIdx = wrappedLines[i].InternalLineFragments[0].StartOriginal;
                    var len = wrappedLines[i].Text.Length;
                    for(int j = startIdx; j< (startIdx + len); j++)
                    {
                        AllChars[j].Line = i;
                    }
                }
            }

            WrappedLineCollection = new TextLineCollection(wrappedLines, InputFragments);
            return WrappedLineCollection;
        }

        /// <summary>
        /// This returns the first letter with input 0,0
        /// The first two with 0,1 etc.
        /// </summary>
        /// <param name="startIdx"></param>
        /// <param name="endIdx"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        private string GetSection(int startIdx, int endIdx)
        {
            if(startIdx > FullTextLastIdx)
            {
                throw new InvalidOperationException($"Cannot GetSection. StartIdx: '{startIdx}' is larger than LastIdx: '{FullTextLastIdx}' ");
            }
            if(endIdx > FullTextLastIdx)
            {
                throw new InvalidOperationException($"Cannot GetSection. EndIdx: '{endIdx}' is larger than LastIdx: '{FullTextLastIdx}'");
            }
            var len = endIdx - startIdx + 1;
            len = len + startIdx > (FullText.Length-1) ? FullTextLength - startIdx : len;
            var subString = FullText.Substring(startIdx, len);
            return subString;
        }

        string GetFullText()
        {
            return FullText;
        }

        public List<string> GetTextOfAllTextRuns()
        {
            List<string> runs = new List<string>();
            foreach (var run in StyleRuns)
            {
                runs.Add(run.Text);
            }
            return runs;
        }

        public int GetParagraphSeparatorCount()
        {
            return ParagraphSeparatorIndicies.Count;
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
