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
using OfficeOpenXml.Drawing.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A serie for a bubble chart
    /// </summary>
    public sealed class ExcelBubbleChartSerie : ExcelChartSerieWithHorizontalErrorBars, IDrawingSerieDataLabel, IDrawingChartDataPoints
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">The chart</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        /// <param name="isPivot">Is pivotchart</param>
        internal ExcelBubbleChartSerie(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot) :
            base(chart,ns, node, isPivot)
        {
            
        }
        ExcelChartSerieDataLabel _dataLabel = null;
        /// <summary>
        /// Datalabel
        /// </summary>
        public ExcelChartSerieDataLabel DataLabel
        {
            get
            {
                if (_dataLabel == null)
                {
                    _dataLabel = new ExcelChartSerieDataLabel(_chart, NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _dataLabel;
            }
        }
        /// <summary>
        /// If the chart has datalabel
        /// </summary>
        public bool HasDataLabel
        {
            get
            {
                return TopNode.SelectSingleNode("c:dLbls", NameSpaceManager) != null;
            }
        }
        const string BUBBLE3D_PATH = "c:bubble3D/@val";
        internal bool Bubble3D
        {
            get
            {
                return GetXmlNodeBool(BUBBLE3D_PATH, true);
            }
            set
            {
                SetXmlNodeBool(BUBBLE3D_PATH, value);    
            }
        }
        const string INVERTIFNEGATIVE_PATH = "c:invertIfNegative/@val";
        internal bool InvertIfNegative
        {
            get
            {
                return GetXmlNodeBool(INVERTIFNEGATIVE_PATH, true);
            }
            set
            {
                SetXmlNodeBool(INVERTIFNEGATIVE_PATH, value);
            }
        }
        /// <summary>
        /// The dataseries for the Bubble Chart
        /// </summary>
        public override string Series
        {
            get
            {
                return base.Series;
            }
            set
            {
                base.Series = value;
                if(string.IsNullOrEmpty(BubbleSize))
                {
                    GenerateLit();
                }
            }
        }        
        const string BUBBLESIZE_TOPPATH = "c:bubbleSize";
        const string BUBBLESIZE_PATH = BUBBLESIZE_TOPPATH + "/c:numRef/c:f";
        /// <summary>
        /// The size of the bubbles
        /// </summary>
        public string BubbleSize
        {
            get
            {
                return GetXmlNodeString(BUBBLESIZE_PATH);
            }
            set
            {
                if(string.IsNullOrEmpty(value))
                {
                    GenerateLit();
                }
                else
                {
                    SetXmlNodeString(BUBBLESIZE_PATH, ExcelCellBase.GetFullAddress(_chart.WorkSheet.Name, value));
                
                    XmlNode cache = TopNode.SelectSingleNode(string.Format("{0}/c:numCache", BUBBLESIZE_PATH), NameSpaceManager);
                    if (cache != null)
                    {
                        cache.ParentNode.RemoveChild(cache);
                    }

                    DeleteNode(string.Format("{0}/c:numLit", BUBBLESIZE_TOPPATH));
                }
            }
        }

        internal void GenerateLit()
        {
            var s = new ExcelAddress(Series);
            var ix = 0;
            var sb = new StringBuilder();
            for (int row = s._fromRow; row <= s._toRow; row++)
            {
                for (int c = s._fromCol; c <= s._toCol; c++)
                {
                    sb.AppendFormat("<c:pt idx=\"{0}\"><c:v>1</c:v></c:pt>", ix++);
                }
            }
            CreateNode(BUBBLESIZE_TOPPATH + "/c:numLit", true);
            XmlNode lit = TopNode.SelectSingleNode(string.Format("{0}/c:numLit", BUBBLESIZE_TOPPATH), NameSpaceManager);
            lit.InnerXml = string.Format("<c:formatCode>General</c:formatCode><c:ptCount val=\"{0}\"/>{1}", ix, sb.ToString());
        }
        ExcelChartDataPointCollection _dataPoints = null;
        /// <summary>
        /// A collection of the individual datapoints
        /// </summary>
        public ExcelChartDataPointCollection DataPoints
        {
            get
            {

                if (_dataPoints == null)
                {
                    _dataPoints = new ExcelChartDataPointCollection(_chart, NameSpaceManager, TopNode, SchemaNodeOrder);
                }
                return _dataPoints;
            }
        }
    }
}
