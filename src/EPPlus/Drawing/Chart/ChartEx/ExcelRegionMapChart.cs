/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2020         EPPlus Software AB           Release EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Packaging;
using System;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Represents a Region Map Chart
    /// </summary>
    public class ExcelRegionMapChart : ExcelChartEx
    {

        internal ExcelRegionMapChart(ExcelDrawings drawings, XmlNode drawingsNode, eChartType? type, XmlDocument chartXml = null, ExcelGroupShape parent = null) :
            base(drawings, drawingsNode, type, chartXml, parent)
        {
            Series.Init(this, NameSpaceManager, TopNode, false, base.Series._list);
            StyleManager.SetChartStyle(Chart.Style.ePresetChartStyle.RegionMapChartStyle1);
        }
        internal ExcelRegionMapChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
            Series.Init(this, NameSpaceManager, TopNode, false, base.Series._list);
        }
        /// <summary>
        /// The series for a region map chart
        /// </summary>
        public new ExcelChartSeries<ExcelRegionMapChartSerie> Series { get; } = new ExcelChartSeries<ExcelRegionMapChartSerie>();
    }
}
