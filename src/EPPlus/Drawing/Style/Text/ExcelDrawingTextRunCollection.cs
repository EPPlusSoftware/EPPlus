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
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    public class ExcelDrawingTextRunCollection : IEnumerable<RegularTextRun>
    {
        ExcelTextFont DefaultRunProperties;

        List<RegularTextRun> textRuns;

        internal ExcelDrawingTextRunCollection()
        {
            textRuns = new List<RegularTextRun>();
        }

        internal ExcelDrawingTextRunCollection(ExcelDrawingParagraph paragraph, ExcelTextFont defaultRunProperties)
        {
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)textRuns).GetEnumerator();
        }

        internal RegularTextRun AddRun(string text)
        {
            var txtRun = new RegularTextRun(text);
            textRuns.Add(txtRun);
            return txtRun;
        }
        internal RegularTextRun Add(RegularTextRun txtRun)
        {
            textRuns.Add(txtRun);
            return txtRun;
        }

        IEnumerator<RegularTextRun> IEnumerable<RegularTextRun>.GetEnumerator()
        {
            return ((IEnumerable<RegularTextRun>)textRuns).GetEnumerator();
        }
    }
}