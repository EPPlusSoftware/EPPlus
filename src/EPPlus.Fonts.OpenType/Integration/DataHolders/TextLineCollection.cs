using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;



namespace EPPlus.Fonts.OpenType.Integration
{
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

        internal MeasurementFont GetFont(int fragIdx)
        {
            return _originalFragments[fragIdx].Font;
        }

        public TextLineCollection(List<TextLineSimple> lines, List<TextFragment> originalFragments)
        {
            _originalFragments = originalFragments;

            //foreach (var line in lines)
            //{
            //    foreach (var lf in line.LineFragments)
            //    {
            //        var text = line.GetLineFragmentText(lf);
            //        smallestTextFragments.Add(text);
            //    }
            //}

            for(int i = 0; i < lines.Count; i++)
            {
                int lineNum = i;
                Lines.Add(new TextLineVizualizer(this, lines[i], i));
                //foreach (var lf in _lines[i].LineFragments)
                //{
                //    var text = _lines[i].GetLineFragmentText(lf);
                //    var font = fonts[lf.RtFragIdx];
                //    var details = _lines[i].LineFragments;
                //    //smallestTextFragments.Add(text);
                //}
            }

            //foreach (var line in _lines)
            //{
            //    //TextLineVizualizer visualizer = ne
            //    foreach (var lf in line.LineFragments)
            //    {
            //        var text = line.GetLineFragmentText(lf);
            //        var font = fonts[lf.RtFragIdx];
            //        int lineNum 
            //        //smallestTextFragments.Add(text);
            //    }
            //    //Lines.Add(new TextLineVizualizer(line));
            //}
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

