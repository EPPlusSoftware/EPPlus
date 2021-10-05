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
using OfficeOpenXml.Style;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// An individual serie item within the chart legend
    /// </summary>
    public class ExcelChartLegendEntry : XmlHelper
    {
        protected ExcelChartStandard _chart;
        internal ExcelChartLegendEntry(XmlNamespaceManager nsm, XmlNode topNode, ExcelChartStandard chart) : base(nsm, topNode)
        {
            Init(chart);
            Index = GetXmlNodeInt("c:idx/@val");
        }

        internal ExcelChartLegendEntry(XmlNamespaceManager nsm, XmlNode legendNode, ExcelChartStandard chart, int serieIndex) : base(nsm)
        {
            Init(chart);
            TopNode = legendNode;
            Index = serieIndex;
            SetXmlNodeInt("c:idx/@val", serieIndex);
        }
        private void Init(ExcelChartStandard chart)
        {
            _chart = chart;
            SchemaNodeOrder = new string[] { "idx", "deleted", "txPr" };
        }
        /// <summary>
        /// The index of the item
        /// </summary>
        public int Index
        {
            get;
            internal set;

        }
        /// <summary>
        /// If the items has been deleted or is visible.
        /// </summary>
        public bool Deleted
        {
            get
            {
                return GetXmlNodeBool("c:delete/@val");
            }
            set
            {
                SetXmlNodeBool("c:delete/@val", value);
            }
        }

        ExcelTextFont _font = null;
        /// <summary>
        /// The Font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                if (_font == null)
                {
                    _font = new ExcelTextFont(_chart, NameSpaceManager, TopNode, $"c:txPr/a:p/a:pPr/a:defRPr", SchemaNodeOrder);
                }
                return _font;
            }
        }
        ExcelTextBody _textBody = null;
        /// <summary>
        /// Access to text body properties
        /// </summary>
        public ExcelTextBody TextBody
        {
            get
            {
                if (_textBody == null)
                {
                    _textBody = new ExcelTextBody(NameSpaceManager, TopNode, $"c:txPr/a:bodyPr", SchemaNodeOrder);
                }
                return _textBody;
            }
        }
        internal void Save()
        {
            if(Deleted==true)
            {
                DeleteNode("c:txPr");
            }
            else
            {
                if (ExistsNode("c:txPr"))
                {
                    DeleteNode("c:delete");
                }
                else
                {
                    TopNode.ParentNode.RemoveChild(TopNode);
                }
            }
        }
    }
}