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
using System.Text;
using System.Xml;
using System.Globalization;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to doughnut chart specific properties
    /// </summary>
    public class ExcelDoughnutChart : ExcelPieChart
    {
        internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot, ExcelGroupShape parent = null) :
            base(drawings, node, type, isPivot, parent)
        {
            
        }
        internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChartBase topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {
        }
        internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
           base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
        }

        internal ExcelDoughnutChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) :
            base(topChart, chartNode, parent)
        {
        }

        string _firstSliceAngPath = "c:firstSliceAng/@val";
        /// <summary>
        /// Angle of the first slize
        /// </summary>
        public decimal FirstSliceAngle
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeDecimal(_firstSliceAngPath);
            }
            internal set
            {
                _chartXmlHelper.SetXmlNodeString(_firstSliceAngPath, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        //string _holeSizePath = "c:chartSpace/c:chart/c:plotArea/{0}/c:holeSize/@val";
        string _holeSizePath = "c:holeSize/@val";
        /// <summary>
        /// Size of the doubnut hole
        /// </summary>
        public decimal HoleSize
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeDecimal(_holeSizePath);
            }
            internal set
            {
                _chartXmlHelper.SetXmlNodeString(_holeSizePath, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        internal override eChartType GetChartType(string name)
        {
            if (name == "doughnutChart")
            {
                if (Series.Count > 0 && Series[0].Explosion > 0)
                {
                    return eChartType.DoughnutExploded;
                }
                else
                {
                    return eChartType.Doughnut;
                }
            }
            return base.GetChartType(name);
        }
    }
}
