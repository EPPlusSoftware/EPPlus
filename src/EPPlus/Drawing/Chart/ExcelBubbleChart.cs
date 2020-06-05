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
using System.Globalization;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Represents a Bar Chart
    /// </summary>
    public sealed class ExcelBubbleChart : ExcelChartStandard, IDrawingDataLabel
    {
        internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent=null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {
            ShowNegativeBubbles = false;
            BubbleScale = 100;
        }

        internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot, ExcelGroupShape parent=null) :
            base(drawings, node, type, isPivot, parent)
        {
        }
        internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent=null) :
            base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
        }
        internal ExcelBubbleChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent=null) :
            base(topChart, chartNode, parent)
        {
        }
        internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            base.InitSeries(chart, ns, node, isPivot, list);
            Series = new ExcelBubbleChartSeries(chart, ns, node, isPivot, base.Series._list);
        }
        string BUBBLESCALE_PATH = "c:bubbleScale/@val";
        /// <summary>
        /// Specifies the scale factor of the bubble chart. Can range from 0 to 300, corresponding to a percentage of the default size,
        /// </summary>
        public int BubbleScale
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeInt(BUBBLESCALE_PATH);
            }
            set
            {
                if(value < 0 && value > 300)
                {
                    throw(new ArgumentOutOfRangeException("Bubblescale out of range. 0-300 allowed"));
                }
                _chartXmlHelper.SetXmlNodeString(BUBBLESCALE_PATH, value.ToString());
            }
        }
        string SHOWNEGBUBBLES_PATH = "c:showNegBubbles/@val";
        /// <summary>
        /// If negative sized bubbles will be shown on a bubble chart
        /// </summary>
        public bool ShowNegativeBubbles
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(SHOWNEGBUBBLES_PATH);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeBool(BUBBLESCALE_PATH, value, true);
            }
        }
        string BUBBLE3D_PATH = "c:bubble3D/@val";
        /// <summary>
        ///If the bubblechart is three dimensional
        /// </summary>
        public bool Bubble3D
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeBool(BUBBLE3D_PATH);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeBool(BUBBLE3D_PATH, value);
                ChartType = value ? eChartType.Bubble3DEffect : eChartType.Bubble;
            }
        }
        string SIZEREPRESENTS_PATH = "c:sizeRepresents/@val";
        /// <summary>
        /// The scale factor for the bubble chart. Can range from 0 to 300, corresponding to a percentage of the default size,
        /// </summary>
        public eSizeRepresents SizeRepresents
        {
            get
            {
                var v = _chartXmlHelper.GetXmlNodeString(SIZEREPRESENTS_PATH).ToLower(CultureInfo.InvariantCulture);
                if (v == "w")
                {
                    return eSizeRepresents.Width;
                }
                else
                {
                    return eSizeRepresents.Area;
                }
            }
            set
            {
                _chartXmlHelper.SetXmlNodeString(SIZEREPRESENTS_PATH, value == eSizeRepresents.Width ? "w" : "area");
            }
        }
        ExcelChartDataLabel _dataLabel = null;
        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (_dataLabel == null)
                {
                    _dataLabel = new ExcelChartDataLabelStandard(Series._chart, NameSpaceManager, ChartNode, "dLbls", _chartXmlHelper.SchemaNodeOrder);
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
                return ChartNode.SelectSingleNode("c:dLbls", NameSpaceManager) != null;
            }
        }

        /// <summary>
        /// The series for a bubble charts
        /// </summary>
        public new ExcelBubbleChartSeries Series { get; private set; }
        internal override eChartType GetChartType(string name)
        {
            if (Bubble3D)
            {
                return eChartType.Bubble3DEffect;
            }
            else
            {
                return eChartType.Bubble;
            }
        }
    }
}
