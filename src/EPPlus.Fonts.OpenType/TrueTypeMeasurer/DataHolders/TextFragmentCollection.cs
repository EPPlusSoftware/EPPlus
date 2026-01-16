using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders
{
    public class TextFragmentCollection
    {
        internal Dictionary<int, CharInfo> CharLookup = new();

        internal List<string> TextFragments = new();

        internal string AllText;
        internal List<int> AllTextNewLineIndicies = new();
        List<TextFragment> _fragmentItems = new List<TextFragment>();
        List<TextLine> _lines = new List<TextLine>();
        List<string> outputFragments = new List<string>();

        public List<int> IndiciesToWrapAt {get; internal set; }
        public TextFragmentCollection(List<string> textFragments)
        {
            TextFragments = textFragments;
            IndiciesToWrapAt = new List<int>();

            //Set data for each fragment and for AllText
            var currentTotalLength = 0;
            for (int i = 0; i < textFragments.Count; i++)
            {
                var currString = textFragments[i];
                var fragment = new TextFragment(currString, currentTotalLength);
                currentTotalLength += currString.Length;
                _fragmentItems.Add(fragment);
            }

            AllText = string.Join(string.Empty, textFragments.ToArray());

            //var allTextLineLess = AllText.Replace("\r", "");
            //allTextLineLess = allTextLineLess.Replace("\n", "");

            //Get the indicies where newlines occur in the combined string
            AllTextNewLineIndicies = GetFirstCharPositionOfNewLines(AllText);

            //Save minor information about each char so each char knows its line/fragment
            int charCount = 0;
            int lineIndex = 0;

            List<int> currFragments = new();
            int lineStartCharIndex = 0;

            //For each fragment
            for (int i = 0; i < TextFragments.Count; i++)
            {
                var textFragment = TextFragments[i];

                //For each char in current fragment
                for (int j = 0; j < textFragment.Length; j++)
                {
                    if (lineIndex <= AllTextNewLineIndicies.Count - 1 &&
                        charCount >= AllTextNewLineIndicies[lineIndex])
                    {
                        var text = AllText.Substring(lineStartCharIndex, charCount - lineStartCharIndex);
                        var trimmedText = text.Trim(['\r', '\n']);

                        var line = new TextLine()
                        {
                            richTextIndicies = currFragments,
                            content = trimmedText,
                            startIndex = lineStartCharIndex,
                        };

                        lineStartCharIndex = AllTextNewLineIndicies[lineIndex];
                        _lines.Add(line);
                        currFragments.Clear();
                        lineIndex++;
                    }

                    var info = new CharInfo(charCount, i, lineIndex);
                    CharLookup.Add(charCount, info);
                    charCount++;
                }
                currFragments.Add(i);
            }

            //Add the last line
            var lastText = AllText.Substring(lineStartCharIndex, charCount - lineStartCharIndex);
            var lastTrimmedText = lastText.Trim(['\r', '\n']);

            var lastLine = new TextLine()
            {
                richTextIndicies = currFragments,
                content = lastTrimmedText,
                startIndex = lineStartCharIndex,
            };

            _lines.Add(lastLine);
            currFragments.Clear();
        }

        public TextFragmentCollection(List<string> textFragments, List<float> fontSizes) : this(textFragments)
        {
            if(textFragments.Count() != fontSizes.Count)
            {
                throw new InvalidOperationException($"TextFragment list and FontSizes list must be equal." +
                    $"Counts:" +
                    $"textFragment: {textFragments.Count()}" +
                    $"fontSizes: {fontSizes.Count()}");
            }

            for(int i = 0; i< textFragments.Count; i++)
            {
                _fragmentItems[i].FontSize = fontSizes[i];
            }
        }

        private List<int> GetFirstCharPositionOfNewLines(string stringsCombined)
        {
            List<int> positions = new List<int>();
            for (int i = 0; i < stringsCombined.Length; i++)
            {
                if (stringsCombined[i] == '\n')
                {
                    if (i < stringsCombined.Length - 1)
                    {
                        positions.Add(i + 1);
                    }
                    else
                    {
                        positions.Add(i);
                    }
                }
            }
            return positions;
        }

        internal void AddWrappingIndex(int index)
        {
            IndiciesToWrapAt.Add(index);
        }

        public List<string> GetFragmentsWithFinalLineBreaks()
        {
            //var newFragments = TextFragments.ToArray();
            foreach(var i in IndiciesToWrapAt)
            {
                _fragmentItems[CharLookup[i].Fragment].AddLocalLbPositon(i);
            }

            foreach(var fragment in _fragmentItems)
            {
                outputFragments.Add(fragment.GetWrappedFragment());
            }

            return outputFragments;
        }

        public List<string> GetFragmentsWithoutLineBreaks()
        {
            //var newFragments = TextFragments.ToArray();
            foreach (var i in IndiciesToWrapAt)
            {
                _fragmentItems[CharLookup[i].Fragment].AddLocalLbPositon(i);
            }

            foreach (var fragment in _fragmentItems)
            {
                outputFragments.Add(fragment.GetWrappedFragment());
            }

            return outputFragments;
        }

        /// <summary>
        /// Should arguably be done in constructor instead of created every time
        /// 
        /// </summary>
        /// <returns></returns>
        public List<float> GetLargestFontSizesOfEachLine()
        {
            List<float> largestFontSizes = new List<float>();
            foreach(var line in _lines)
            {
                float largest = float.MinValue;
                foreach(var idx in line.richTextIndicies)
                {
                    if (float.IsNaN(_fragmentItems[idx].FontSize) == false && _fragmentItems[idx].FontSize > largest)
                    {
                        largest = _fragmentItems[idx].FontSize;
                    }
                }
                if (largest != float.MinValue)
                {
                    largestFontSizes.Add(largest);
                }
            }
            return largestFontSizes;
        }

        public List<int> GetLinesFragmentIsPartOf(int fragmentIndex)
        {
            //Add throw if index does not exist?
            return _fragmentItems[fragmentIndex].IsPartOfLines;
        }
    }
}
