using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.RichText;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;



namespace EPPlus.Fonts.OpenType.Integration
{

    public class TextLineCollection : List<TextLineSimple>, IEnumerable<TextLineSimple>
    {
        /// <summary>
        /// Array of the line numbers where fragment occurs
        /// </summary>
        public List<LineFragmentOutput> LineFragments = new List<LineFragmentOutput>();
        List<ITextFragmentBase> _originalFragments;

        public int[] GetLineNumbersThatUse(TextFragment fragment)
        {
            var idx = _originalFragments.IndexOf(fragment);
            List<int> result = new List<int>();
            if (idx != -1)
            {
                foreach (var key in fragIdLookup[idx].Keys)
                {
                    result.Add(key);
                }
                return result.ToArray();
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Returns null if fragment is not found in any lines
        /// </summary>
        /// <param name="fragment"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public List<TextLineSimple> GetTextLinesThatUse(TextFragment fragment)
        {
            var idx = _originalFragments.IndexOf(fragment);

            if(idx != -1)
            {
                List<TextLineSimple> retLines = new List<TextLineSimple>();
                foreach (var key in fragIdLookup[idx].Keys)
                {
                    retLines.Add(this[key]);
                }
                return retLines;
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// Returns null if fragment is not found in any linefragments
        /// </summary>
        /// <param name="fragment"></param>
        /// <returns></returns>
        public List<LineFragment> GetInternalLineFragmentsThatUse(TextFragment fragment)
        {
            var idx = _originalFragments.IndexOf(fragment);

            List<LineFragment> retFragments = null;

            if (idx != -1)
            {
                retFragments = new List<LineFragment>();

                foreach (var key in fragIdLookup[idx].Keys)
                {
                    foreach(var lineFragment in fragIdLookup[idx][key])
                    {
                        retFragments.Add(this[key].InternalLineFragments[lineFragment]);
                    }
                }
            }

            return retFragments;
        }

        /// <summary>
        /// Returns null if fragment is not found in any linefragments
        /// </summary>
        /// <param name="fragment"></param>
        /// <returns></returns>
        public List<LineFragmentOutput> GetLineFragmentsThatUse(TextFragment fragment)
        {
            var idx = _originalFragments.IndexOf(fragment);

            List<LineFragmentOutput> retFragments = null;

            if (idx != -1)
            {
                retFragments = new List<LineFragmentOutput>();

                foreach (var key in fragIdLookup[idx].Keys)
                {
                    foreach (var lineFragment in fragIdLookup[idx][key])
                    {
                        retFragments.Add(this[key].LineFragments[lineFragment]);
                    }
                }
            }

            return retFragments;
        }

        /// <summary>
        /// The id of the orginal fragment may correspond to 
        /// multiple lines with multiple different richtext fragments
        /// So. fragIdLookup[fragId] returns dictionary of lines that contains the fragment
        /// fragIdLookup[fragId][lineNum] returns list of output fragments that contain the font
        /// fragIdLookup[fragId][lineNum][0] returns first richtextfragment within the line that contains the font.
        /// </summary>
        Dictionary<int, Dictionary<int, List<int>>> fragIdLookup = new Dictionary<int, Dictionary<int, List<int>>>();

        //internal MeasurementFont GetFont(int fragIdx)
        //{
        //    return _originalFragments[fragIdx].RichTextOptions;
        //}

        /// <summary>
        /// If using this MUST call FinalizeTextLineData to finish the information gathering
        /// </summary>
        /// <param name="fragmentCollection"></param>
        public TextLineCollection(TextFragmentCollectionSimple fragmentCollection)
        {
            _originalFragments = fragmentCollection;

            for (int i = 0; i < fragmentCollection.Count; i++)
            {
                fragIdLookup.Add(i, new Dictionary<int, List<int>>());
            }
        }

        private void AddToDictionary(int idx, int lineNum, int fragPosInline)
        {

            if (fragIdLookup[idx].ContainsKey(lineNum) == false)
            {
                fragIdLookup[idx].Add(lineNum, new List<int>());
            }

            fragIdLookup[idx][lineNum].Add(fragPosInline);
        }

        internal void FinalizeTextLineData(List<TextLineSimple> lines)
        {
            for (int i = 0; i < lines.Count; i++)
            {
                int lineNum = i;
                int fragCount = 0;

                foreach (var lf in lines[i].InternalLineFragments)
                {
                    var idx = lf.FragmentIndex;
                    AddToDictionary(idx, lineNum, fragCount);
                    LineFragmentOutput data;
                    if (lines[i].LineFragments == null || lines[i].LineFragments.Count != lines[i].InternalLineFragments.Count)
                    {
                        data = new LineFragmentOutput(
                           () => { return _originalFragments[idx]; },
                           () => { return lf.Width; },
                           () => { return lf.StartIdx; },
                           () => { return lf.StartRt; },
                           () => { return lf.StartOriginal; },
                           lines[i].GetLineFragmentText(lf)
                           );
                    }
                    else
                    {
                        //If already calculate don't duplicate just use the data
                        data = lines[i].LineFragments[fragCount];
                    }

                    LineFragments.Add(data);

                    fragCount++;
                }
                Add(lines[i]);
            }
        }

        public double LargestWidthWithSpace { get; private set; } = -1d;
        public double LargestWidthWithoutSpace { get; private set; } = -1d;
        public int idxOfLargestLine { get; private set; } = -1;
        public List<double> SpaceWidthsPerLine { get; private set; } = new List<double>();

        public TextLineCollection(List<TextLineSimple> lines, List<ITextFragmentBase> originalFragments)
        {
            int lineIdx = 0;
            foreach (var line in lines)
            {
                double largestAscent = 0;
                double largestDescent = 0;
                double largestFontSize = 0;
                foreach (var lineFragment in line.InternalLineFragments)
                {
                    var frag = originalFragments[lineFragment.FragmentIndex];
                    if (frag == null) continue;
                    largestAscent = Math.Max(frag.AscentPoints, largestAscent);
                    largestDescent = Math.Max(frag.DescentPoints, largestDescent);
                    largestFontSize = Math.Max(largestFontSize, frag.Size);
                }
                line.LargestAscent = largestAscent;
                line.LargestDescent = largestDescent;
                line.LargestFontSize = largestFontSize;

                line.FinalizeLineFragments(originalFragments);

                if(line.Width > LargestWidthWithSpace)
                {
                    LargestWidthWithSpace = line.Width;

                }
                if(line.GetWidthWithoutTrailingSpaces() > LargestWidthWithoutSpace)
                {
                    LargestWidthWithoutSpace = line.GetWidthWithoutTrailingSpaces();
                    idxOfLargestLine = lineIdx;
                }
                LargestWidthWithSpace = Math.Max(LargestWidthWithSpace, line.Width);
                LargestWidthWithoutSpace = Math.Max(LargestWidthWithoutSpace, line.GetWidthWithoutTrailingSpaces());
                SpaceWidthsPerLine.Add(line.lastFontSpaceWidth);
                lineIdx++;
            }

            _originalFragments = originalFragments;

            for(int i = 0; i < originalFragments.Count; i++)
            {
                fragIdLookup.Add(i, new Dictionary<int, List<int>>());
            }

            FinalizeTextLineData(lines);
        }
    }
}

