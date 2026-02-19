/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Represents an individual datalabel
    /// </summary>
    public class ExcelChartDataLabelItem : ExcelChartDataLabelStandard
    {
        string _fontPropertiesPath = "";
        internal ExcelChartDataLabelItem(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, string nodeName, string[] schemaNodeOrder)
           : base(chart, ns, node, nodeName, schemaNodeOrder)
        {
            Layout = new ExcelLayout(NameSpaceManager, TopNode, $"c:layout","c:extLst/c:ext[1]/c15:layout",  SchemaNodeOrder);
            _fontPropertiesPath = $"{NsPrefix}:tx/{NsPrefix}:rich";
        }

        /// <summary>
        /// Define position for manual elements
        /// </summary>
        public ExcelLayout Layout { get; private set; }

        ExcelParagraphCollection _paragraphs = null;

        /// <summary>
        /// Access to text body properties
        /// </summary>
        private ExcelParagraphCollection ParagraphCollection
        {
            get
            {
                if (_paragraphs == null)
                {
                    //var firstParaPath = _textBodyPropertiesParentPath + $"/{NsPrefix}:p";
                    //par.SelectNodes("a:r", NameSpaceManager);
                    _paragraphs = new ExcelParagraphCollection(_chart, NameSpaceManager, TopNode, _fontPropertiesPath + "/a:p", SchemaNodeOrder);
                }
                return _paragraphs;
            }
        }

        /// <summary>
        /// Replace datalabel text
        /// </summary>
        /// <param name="replacementText"></param>
        public void SetText(string replacementText)
        {
            ParagraphCollection.Clear();
            ParagraphCollection.Add(replacementText, true);
        }

        internal void AddField(string fldType)
        {
            //Only add if none exist
            var addParagraph = ParagraphCollection.Count == 0;

            ParagraphCollection.AddFieldNode(fldType, addParagraph);
        }

        internal List<List<string>> GetExistingParagraphStrings()
        {
            return ParagraphCollection.GetParagraphTextLists();
        }
        /// <summary>
        /// The index of an individual datalabel
        /// </summary>
        public int Index
        {
            get
            {
                return GetXmlNodeInt("c:idx/@val");
            }
            set
            {
                SetXmlNodeString("c:idx/@val", value.ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}
