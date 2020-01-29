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
using OfficeOpenXml.Table.PivotTable;
using System.Globalization;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Provides access to OfPie-Chart specific properties
    /// </summary>
    public class ExcelOfPieChart : ExcelPieChart
    {
        internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot, ExcelGroupShape parent = null) :
            base(drawings, node, type, isPivot, parent)
        {
                SetTypeProperties();
        }
        internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, eChartType? type, ExcelChart topChart, ExcelPivotTable PivotTableSource, XmlDocument chartXml, ExcelGroupShape parent = null) :
            base(drawings, node, type, topChart, PivotTableSource, chartXml, parent)
        {
            SetTypeProperties();
        }

        internal ExcelOfPieChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode, ExcelGroupShape parent = null) :
           base(drawings, node, uriChart, part, chartXml, chartNode, parent)
        {
            SetTypeProperties();
        }
        internal ExcelOfPieChart(ExcelChart topChart, XmlNode chartNode, ExcelGroupShape parent = null) :
           base(topChart, chartNode, parent)
        {
            SetTypeProperties();
        }

        private void SetTypeProperties()
        {
            if (ChartType == eChartType.BarOfPie)
            {
                OfPieType = ePieType.Bar;
            }
            else
            {
                OfPieType = ePieType.Pie;
            }
        }

        const string pieTypePath = "c:ofPieType/@val";
        /// <summary>
        /// Type, pie or bar
        /// </summary>
        public ePieType OfPieType
        {
            get
            {
                if (_chartXmlHelper.GetXmlNodeString(pieTypePath) == "bar")
                    return ePieType.Bar;
                else
                {
                    return ePieType.Pie;
                }
            }
            internal set
            {
                _chartXmlHelper.CreateNode(pieTypePath, true);
                _chartXmlHelper.SetXmlNodeString(pieTypePath, value == ePieType.Bar ? "bar" : "pie");
            }
        }

        readonly string _gapWidthPath = "c:gapWidth/@val";
        /// <summary>
        /// The size of the gap between two adjacent bars/columns
        /// </summary>
        public int GapWidth
        {
            get
            {
                return _chartXmlHelper.GetXmlNodeInt(_gapWidthPath);
            }
            set
            {
                _chartXmlHelper.SetXmlNodeString(_gapWidthPath, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        internal override eChartType GetChartType(string name)
        {
            if (name == "ofPieChart")
            {
                if (OfPieType==ePieType.Bar)
                {
                    return eChartType.BarOfPie;
                }
                else
                {
                    return eChartType.PieOfPie;
                }
            }
            return base.GetChartType(name);
        }
    }
}
