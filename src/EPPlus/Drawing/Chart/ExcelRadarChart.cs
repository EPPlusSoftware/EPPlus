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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to line chart specific properties
    /// </summary>
    public class ExcelRadarChart : ExcelChartStandard, IDrawingDataLabel
    {
        #region "Constructors"
        internal ExcelRadarChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
            SetTypeProperties();
        }

        internal ExcelRadarChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(topChart, chartNode, parent)
        {
            SetTypeProperties();
        }
        internal ExcelRadarChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {
            SetTypeProperties();
        }
        #endregion
        internal override void InitSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list = null)
        {
            base.InitSeries(chart, ns, node, isPivot, list);
            Series.Init(chart, ns, node, isPivot, base.Series._list);
        }
        private void SetTypeProperties()
        {
            if (ChartType == eChartType.RadarFilled)
            {
                RadarStyle = eRadarStyle.Filled;
            }
            else if  (ChartType == eChartType.RadarMarkers)
            {
                RadarStyle =  eRadarStyle.Marker;
            }
            else
            {
                RadarStyle = eRadarStyle.Standard;
            }
        }

        string STYLE_PATH = "c:radarStyle/@val";
        /// <summary>
        /// The type of radarchart
        /// </summary>
        public eRadarStyle RadarStyle
        {
            get
            {
                var v=_chartXmlHelper.GetXmlNodeString(STYLE_PATH);
                if (string.IsNullOrEmpty(v))
                {
                    return eRadarStyle.Standard;
                }
                else
                {
                    return (eRadarStyle)Enum.Parse(typeof(eRadarStyle), v, true);
                }
            }
            set
            {
                _chartXmlHelper.SetXmlNodeString(STYLE_PATH, value.ToString().ToLower(CultureInfo.InvariantCulture));
            }
        }
        ExcelChartDataLabel _DataLabel = null;
        /// <summary>
        /// Access to datalabel properties
        /// </summary>
        public ExcelChartDataLabel DataLabel
        {
            get
            {
                if (_DataLabel == null)
                {
                    _DataLabel = new ExcelChartDataLabelStandard(this, NameSpaceManager, ChartNode, "dLbls", _chartXmlHelper.SchemaNodeOrder);
                }
                return _DataLabel;
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
        internal override eChartType GetChartType(string name)
        {
            if (RadarStyle == eRadarStyle.Filled)
            {
                return eChartType.RadarFilled;
            }
            else if (RadarStyle == eRadarStyle.Marker)
            {
                return eChartType.RadarMarkers;
            }
            else
            {
                return eChartType.Radar;
            }
        }
        /// <summary>
        /// A collection of series for a Radar Chart
        /// </summary>
        public new ExcelChartSeries<ExcelRadarChartSerie> Series { get; } = new ExcelChartSeries<ExcelRadarChartSerie>();
    }
}
