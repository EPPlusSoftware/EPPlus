//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Xml;

//namespace EPPlus.Fonts.OpenType.Integration
//{
//    public class TextFragmentCollection
//    {
//        internal Dictionary<int, CharInfo> CharLookup = new();

//        internal List<string> TextFragments = new();

//        internal string AllText;
//        internal List<int> AllTextNewLineIndicies = new();
//        List<TextFragment> _fragmentItems = new List<TextFragment>();
//        List<TextLine> _lines = new List<TextLine>();
//        List<string> outputFragments = new List<string>();
//        //List<string> outputStrings = new List<string>();
//        internal List<double> LargestFontSizePerLine { get; private set; } = new();
//        internal List<double> AscentPerLine { get; private set; } = new();
//        internal List<double> DescentPerLine { get; private set; } = new();

//        /// <summary>
//        /// Added linewidths in points
//        /// </summary>
//        internal List<double> lineWidths { get; private set; } = new List<double>();

//        public List<int> IndiciesToWrapAt {get; internal set; }

//        List<TextLineSimple> TextLines = new();

//        public TextFragmentCollection(List<string> textFragments)
//        {
//            TextFragments = textFragments;
//            IndiciesToWrapAt = new List<int>();

//            //Set data for each fragment and for AllText
//            var currentTotalLength = 0;
//            for (int i = 0; i < textFragments.Count; i++)
//            {
//                var currString = textFragments[i];
//                var fragment = new TextFragment(currString, currentTotalLength);
//                currentTotalLength += currString.Length;
//                _fragmentItems.Add(fragment);
//            }

//            AllText = string.Join(string.Empty, textFragments.ToArray());

//            //Get the indicies where newlines occur in the combined string
//            AllTextNewLineIndicies = GetFirstCharPositionOfNewLines(AllText);

//            //Save minor information about each char so each char knows its line/fragment
//            int charCount = 0;
//            int lineIndex = 0;

//            List<int> currFragments = new();
//            int lineStartCharIndex = 0;

//            //For each fragment
//            for (int i = 0; i < TextFragments.Count; i++)
//            {
//                var textFragment = TextFragments[i];

//                //For each char in current fragment
//                for (int j = 0; j < textFragment.Length; j++)
//                {
//                    if (lineIndex <= AllTextNewLineIndicies.Count - 1 &&
//                        charCount >= AllTextNewLineIndicies[lineIndex])
//                    {
//                        var text = AllText.Substring(lineStartCharIndex, charCount - lineStartCharIndex);
//                        var trimmedText = text.Trim(['\r', '\n']);

//                        var line = new TextLine()
//                        {
//                            richTextIndicies = currFragments,
//                            content = trimmedText,
//                            startIndex = lineStartCharIndex,
//                        };

//                        lineStartCharIndex = AllTextNewLineIndicies[lineIndex];
//                        _lines.Add(line);
//                        currFragments.Clear();
//                        lineIndex++;
//                    }

//                    var info = new CharInfo(charCount, i, lineIndex);
//                    CharLookup.Add(charCount, info);
//                    charCount++;
//                }
//                currFragments.Add(i);
//            }

//            //Add the last line
//            var lastText = AllText.Substring(lineStartCharIndex, charCount - lineStartCharIndex);
//            var lastTrimmedText = lastText.Trim(['\r', '\n']);

//            var lastLine = new TextLine()
//            {
//                richTextIndicies = currFragments,
//                content = lastTrimmedText,
//                startIndex = lineStartCharIndex,
//            };

//            _lines.Add(lastLine);
//            currFragments.Clear();
//        }

//        public TextFragmentCollection(List<string> textFragments, List<float> fontSizes) : this(textFragments)
//        {
//            if(textFragments.Count() != fontSizes.Count)
//            {
//                throw new InvalidOperationException($"TextFragment list and FontSizes list must be equal." +
//                    $"Counts:" +
//                    $"textFragment: {textFragments.Count()}" +
//                    $"fontSizes: {fontSizes.Count()}");
//            }

//            for(int i = 0; i< textFragments.Count; i++)
//            {
//                _fragmentItems[i].FontSize = fontSizes[i];
//            }
//        }

//        private List<int> GetFirstCharPositionOfNewLines(string stringsCombined)
//        {
//            List<int> positions = new List<int>();
//            for (int i = 0; i < stringsCombined.Length; i++)
//            {
//                if (stringsCombined[i] == '\n')
//                {
//                    if (i < stringsCombined.Length - 1)
//                    {
//                        positions.Add(i + 1);
//                    }
//                    else
//                    {
//                        positions.Add(i);
//                    }
//                }
//            }
//            return positions;
//        }

//        internal void AddLargestFontSizePerLine(double fontSize)
//        {
//            LargestFontSizePerLine.Add(fontSize);
//        }

//        internal void AddAscentPerLine(double Ascent)
//        {
//            AscentPerLine.Add(Ascent);
//        }

//        internal void AddDescentPerLine(double Descent)
//        {
//            DescentPerLine.Add(Descent);
//        }

//        internal void AddWrappingIndex(int index)
//        {
//            IndiciesToWrapAt.Add(index);
//        }

//        public List<string> GetFragmentsWithFinalLineBreaks()
//        {
//            //var newFragments = TextFragments.ToArray();
//            foreach(var i in IndiciesToWrapAt)
//            {
//                _fragmentItems[CharLookup[i].Fragment].AddLocalLbPositon(i);
//            }

//            foreach(var fragment in _fragmentItems)
//            {
//                outputFragments.Add(fragment.GetWrappedFragment());
//            }

//            return outputFragments;
//        }

//        public List<TextLine> GetOutputLines()
//        {
//            var fragments = GetFragmentsWithFinalLineBreaks();

//            string outputAllText = "";

//            //Create final all text
//            for (int i = 0; i < fragments.Count(); i++)
//            {
//                outputAllText += fragments[i];
//            }

//            var newLineIndiciesOutput = GetFirstCharPositionOfNewLines(outputAllText);
//            List<TextLine> outputTextLines = new List<TextLine>();

//            //Save minor information about each char so each char knows its line/fragment
//            int charCount = 0;
//            int lineIndex = 0;

//            List<int> currFragments = new();
//            int lineStartCharIndex = 0;
//            int RtInternalStartIndex = 0;

//            //For each fragment
//            for (int i = 0; i < fragments.Count; i++)
//            {
//                var textFragment = fragments[i];

//                //For each char in current fragment
//                for (int j = 0; j < textFragment.Length; j++)
//                {
//                    if (lineIndex <= newLineIndiciesOutput.Count - 1 &&
//                        charCount >= newLineIndiciesOutput[lineIndex])
//                    {
//                        //We have reached a new line
//                        var text = outputAllText.Substring(lineStartCharIndex, charCount - lineStartCharIndex);
//                        var trimmedText = text.Trim(['\r', '\n']);

//                        var line = new TextLine()
//                        {
//                            richTextIndicies = currFragments,
//                            content = trimmedText,
//                            startIndex = lineStartCharIndex,
//                            rtContentStartIndexPerRt = new List<int>(j),
//                            lastRtInternalIndex = j,
//                            startRtInternalIndex = RtInternalStartIndex
//                        };

//                        lineStartCharIndex = newLineIndiciesOutput[lineIndex];
//                        outputTextLines.Add(line);
//                        currFragments.Clear();
//                        RtInternalStartIndex = j;
//                        lineIndex++;
//                    }

//                    //var info = new CharInfo(charCount, i, lineIndex);
//                    //CharLookup.Add(charCount, info);
//                    charCount++;
//                }
//                currFragments.Add(i);
//            }

//            //Add the last line
//            var lastText = outputAllText.Substring(lineStartCharIndex, charCount - lineStartCharIndex);
//            var lastTrimmedText = lastText.Trim(['\r', '\n']);

//            var lastLine = new TextLine()
//            {
//                richTextIndicies = currFragments,
//                content = lastTrimmedText,
//                startIndex = lineStartCharIndex,
//                lastRtInternalIndex = fragments.Last().Length,
//                startRtInternalIndex = RtInternalStartIndex
//            };

//            outputTextLines.Add(lastLine);
//            currFragments.Clear();

//            return outputTextLines;
//        }

//        //public void GetTextFragmentOutputStartIndiciesOnFinalLines(List<string> finalOutputLines)
//        //{
//        //    var lineBreakStrings = GetFragmentsWithFinalLineBreaks();
//        //    //var fragmentsWithoutLineBreaks = GetFragmentsWithoutLineBreaks();
//        //    string outputAllText = "";
            
//        //    //List<List<int>> eachFragmentForEachLine = new List<List<int>>();

//        //    for(int i = 0; i< finalOutputLines.Count(); i++)
//        //    {
//        //        var line = new TextLine();
//        //        line.content = finalOutputLines[i];

//        //    }

//        //    //Create final all text
//        //    for(int i = 0; i< lineBreakStrings.Count(); i++)
//        //    {

//        //        ////var length = lineBreakStrings[i].Length;
//        //        //outputAllText += lineBreakStrings[i];
//        //    }


//        //}

//        //public List<TextLineSimple> GetSimpleTextLines(List<string> wrappedLines)
//        //{
//        //    //The wrapped lines
//        //    var lines = wrappedLines;
//        //    //The richText data in a form that is one to one in chars
//        //    var frags = GetFragmentsWithoutLineBreaks();

//        //    int fragIdx = 0;
//        //    int charIdx = 0;
//        //    int lastFragLength = 0;

//        //    var currentFragment = frags[fragIdx];

//        //    for (int i = 0; i< wrappedLines.Count(); i++)
//        //    {
//        //        var textLineSimple = new TextLineSimple();

//        //        for (int j = 0; j < wrappedLines[i].Length; j++)
//        //        {
//        //            textLineSimple
//        //            charIdx++;
//        //        }
//        //    }
//        //}

//        public List<string> GetFragmentsWithoutLineBreaks()
//        {
//            ////var newFragments = TextFragments.ToArray();
//            //foreach (var i in IndiciesToWrapAt)
//            //{
//            //    _fragmentItems[CharLookup[i].Fragment].AddLocalLbPositon(i);
//            //}

//            foreach (var fragment in _fragmentItems)
//            {
//                outputFragments.Add(fragment.GetLineBreakFreeFragment());
//            }

//            return outputFragments;
//        }

//        internal void AddLineWidth(double width)
//        {
//            lineWidths.Add(width);
//        }

//        /// <summary>
//        /// Should arguably be done in constructor instead of created every time
//        /// 
//        /// </summary>
//        /// <returns></returns>
//        public List<double> GetLargestFontSizesOfEachLine()
//        {
//            //List<float> largestFontSizes = new List<float>();
//            //foreach(var line in _lines)
//            //{
//            //    float largest = float.MinValue;
//            //    foreach(var idx in line.richTextIndicies)
//            //    {
//            //        if (float.IsNaN(_fragmentItems[idx].FontSize) == false && _fragmentItems[idx].FontSize > largest)
//            //        {
//            //            largest = _fragmentItems[idx].FontSize;
//            //        }
//            //    }
//            //    if (largest != float.MinValue)
//            //    {
//            //        largestFontSizes.Add(largest);
//            //    }
//            //}
//            //return largestFontSizes;
//            return LargestFontSizePerLine;
//        }

//        public double GetAscent(int lineIdx)
//        {
//            return AscentPerLine[lineIdx];
//        }

//        public double GetDescent(int lineIdx)
//        {
//            return DescentPerLine[lineIdx];
//        }

//        internal void AddFragmentWidth(int fragmentIdx,  double width)
//        {
//            _fragmentItems[fragmentIdx].PointWidthPerOutputLineInFragment.Add(width);
//        }

//        public List<double> GetFragmentWidths(int fragmentIdx)
//        {
//            return _fragmentItems[fragmentIdx].PointWidthPerOutputLineInFragment;
//        }

//        public List<int> GetLinesFragmentIsPartOf(int fragmentIndex)
//        {
//            //Add throw if index does not exist?
//            return _fragmentItems[fragmentIndex].IsPartOfLines;
//        }
//    }
//}
