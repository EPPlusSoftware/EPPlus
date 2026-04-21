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

        public TextLineVizualizer(TextLineCollection parentCollection, TextLineSimple line, int lineNum)
        {
            _lineNum = lineNum;
            TextLineDetails = line;
            _parentCollection = parentCollection;

            foreach(var lf in line.LineFragments)
            {
                lineFragmentText.Add(line.GetLineFragmentText(lf));
                fragIds.Add(lf.RtFragIdx);
            }
            //var text = _lines[i].GetLineFragmentText(lf);
            //var font = fonts[lf.RtFragIdx];
            //var details = _lines[i].LineFragments;

        }
    }
}
