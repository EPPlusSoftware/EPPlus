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
        List<TextFragment> fragmentItems = new List<TextFragment>();

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
                fragmentItems.Add(fragment);
            }

            AllText = string.Join(string.Empty, textFragments.ToArray());
            //Get the indicies where newlines occur in the combined string
            AllTextNewLineIndicies = GetFirstCharPositionOfNewLines(AllText);

            //Save minor information about each char so each char knows its line/fragment
            int charCount = 0;
            int lineIndex = 0;

            List<int> currFragments = new();
            List<TextLine> lines = new List<TextLine>();
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
                        lines.Add(line);
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

            lines.Add(lastLine);
            currFragments.Clear();
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
                fragmentItems[CharLookup[i].Fragment].AddLocalLbPositon(i);
            }

            List<string> finalFragments = new List<string>();

            foreach(var fragment in fragmentItems)
            {
                finalFragments.Add(fragment.GetWrappedFragment());
            }

            return finalFragments;
        }
    }
}
