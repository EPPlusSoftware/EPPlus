/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    9/11/2025         EPPlus Software AB       EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelDrawingTextRunCollection : XmlHelper, IEnumerable<ExcelParagraphTextRun>
    {
        List<ExcelParagraphTextRun> _textRuns;
        ExcelDrawingParagraph _paragraph;
        Action _initXml;
        internal ExcelDrawingTextRunCollection(ExcelDrawingParagraph paragraph, XmlNamespaceManager nsm, XmlNode topNode, Action initXml) : base(nsm, topNode)
        {
            _paragraph = paragraph;
            SchemaNodeOrder = _paragraph.SchemaNodeOrder;
            _initXml = initXml;
            _textRuns = new List<ExcelParagraphTextRun>();            
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)_textRuns).GetEnumerator();
        }

        public ExcelParagraphTextRun Add(string text)
        {            
            var txtRun = new ExcelParagraphTextRun(_paragraph._prd, NameSpaceManager, TopNode);
            _textRuns.Add(txtRun);
            return txtRun;
        }
        internal ExcelParagraphTextRun Add(ExcelParagraphTextRun txtRun)
        {
            _textRuns.Add(txtRun);
            return txtRun;
        }

        IEnumerator<ExcelParagraphTextRun> IEnumerable<ExcelParagraphTextRun>.GetEnumerator()
        {
            return ((IEnumerable<ExcelParagraphTextRun>)_textRuns).GetEnumerator();
        }
    }
}