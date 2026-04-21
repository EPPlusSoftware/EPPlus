using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;



namespace EPPlus.Fonts.OpenType.Integration
{
    [DebuggerDisplay("Lines = {Lines}")]
    public class TextLineCollection : List<TextLineSimple>, IEnumerable<TextLineSimple>
    {

        public List<TextFragment> GetFragments()
        {
            List<TextFragment> fragments = new List<TextFragment>();

            for (int i = 0; i< Lines.Count; i++)
            {
                var line = Lines[i];
                for (int j = 0; j< line.lineFragmentText.Count; j++)
                {
                    var fragment = new TextFragment() 
                    { 
                        Text = line.lineFragmentText[j],
                        Font = GetFont(line.fragIds[j]),
                    };
                    fragments.Add(fragment);
                }
            }

            return fragments;
        }

        internal List<TextLineVizualizer> Lines = new List<TextLineVizualizer>();

        List<TextFragment> _originalFragments;

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

                Lines.Add(new TextLineVizualizer(this, lines[i], i, ref fragIdLookup));
            }
        }

        IEnumerator<TextLineSimple> IEnumerable<TextLineSimple>.GetEnumerator()
        {
            for (int i = 0; i < Lines.Count; i++)
            {
                yield return Lines[i].TextLineDetails;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return Lines.GetEnumerator();
        }
    }
}

