using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration
{

    public class TextLineVizualizer
    {
        TextLineCollection _parentCollection;
        internal TextLineSimple TextLineDetails;
        int _lineNum;

        internal List<string> lineFragmentText = new List<string>();
        internal List<int> fragIds = new List<int>();

        //internal MeasurementFont font 
        //{ get 
        //    { 
        //       return _parentCollection.GetFont(TextLineDetails.)
        //    } 
        //}

        public TextLineVizualizer(TextLineCollection parentCollection, TextLineSimple line, int lineNum, 
            ref Dictionary<int, Dictionary<int, List<int>>> fragmentLookup)
        {
            _lineNum = lineNum;
            TextLineDetails = line;
            _parentCollection = parentCollection;
            int fragCount = 0;
            foreach(var lf in line.LineFragments)
            {
                lineFragmentText.Add(line.GetLineFragmentText(lf));
                fragIds.Add(lf.RtFragIdx);

                if (fragmentLookup[lf.RtFragIdx].ContainsKey(lineNum)== false)
                {
                    fragmentLookup[lf.RtFragIdx].Add(lineNum, new List<int>());
                }

                fragmentLookup[lf.RtFragIdx][lineNum].Add(fragCount);

                fragCount++;
            }
            //var text = _lines[i].GetLineFragmentText(lf);
            //var font = fonts[lf.RtFragIdx];
            //var details = _lines[i].LineFragments;

        }
    }
}
