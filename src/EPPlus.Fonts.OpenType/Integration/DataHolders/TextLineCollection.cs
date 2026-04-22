using OfficeOpenXml.Interfaces.Drawing.Text;
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
        public List<LineFragmentData> LineFragments = new List<LineFragmentData>();
        List<TextFragment> _originalFragments;

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
        public List<LineFragment> GetLineFragmentsThatUse(TextFragment fragment)
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

        ///// <summary>
        ///// Returns null if fragment is not found in any linefragments
        ///// </summary>
        ///// <param name="fragment"></param>
        ///// <returns></returns>
        //public List<LineFragment> GetLineFragmentDataThatUses(TextFragment fragment)
        //{
        //    var idx = _originalFragments.IndexOf(fragment);

        //    List<LineFragmentData> retFragments = null;

        //    if (idx != -1)
        //    {
        //        retFragments = new List<LineFragmentData>();

        //        foreach (var key in fragIdLookup[idx].Keys)
        //        {
        //            foreach (var lineFragment in fragIdLookup[idx][key])
        //            {
        //                retFragments.Add(this[key].LineFragments[lineFragment]);
        //            }
        //        }
        //    }

        //    return retFragments;
        //}

        /// <summary>
        /// The id of the orginal fragment may correspond to 
        /// multiple lines with multiple different richtext fragments
        /// So. fragIdLookup[fragId] returns dictionary of lines that contains the fragment
        /// fragIdLookup[fragId][lineNum] returns list of output fragments that contain the font
        /// fragIdLookup[fragId][lineNum][0] returns first richtextfragment within the line that contains the font.
        /// </summary>
        Dictionary<int, Dictionary<int, List<int>>> fragIdLookup = new Dictionary<int, Dictionary<int, List<int>>>();

        internal MeasurementFont GetFont(int fragIdx)
        {
            return _originalFragments[fragIdx].Font;
        }

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

                    LineFragmentData data = new LineFragmentData(
                        () => { return _originalFragments[idx]; },
                        () => { return lf.Width; },
                        lines[i].GetLineFragmentText(lf));
                    LineFragments.Add(data);

                    fragCount++;
                }
                Add(lines[i]);
            }
        }

        public TextLineCollection(List<TextLineSimple> lines, List<TextFragment> originalFragments)
        {
            _originalFragments = originalFragments;

            for(int i = 0; i < originalFragments.Count; i++)
            {
                fragIdLookup.Add(i, new Dictionary<int, List<int>>());
            }

            for(int i = 0; i < lines.Count; i++)
            {
                int lineNum = i;
                int fragCount = 0;

                lines[i].CreateFinalizedSubstringsInLineFragments();

                foreach (var lf in lines[i].InternalLineFragments)
                {
                    var idx = lf.FragmentIndex;

                    if (fragIdLookup[idx].ContainsKey(lineNum) == false)
                    {
                        fragIdLookup[idx].Add(lineNum, new List<int>());
                    }

                    fragIdLookup[idx][lineNum].Add(fragCount);

                    LineFragmentData data = new LineFragmentData(
                        () => { return _originalFragments[idx]; },
                        () => { return lf.Width; },
                        lines[i].GetLineFragmentText(lf));
                    LineFragments.Add(data);

                    fragCount++;
                }
                Add(lines[i]);
            }
        }
    }
}

