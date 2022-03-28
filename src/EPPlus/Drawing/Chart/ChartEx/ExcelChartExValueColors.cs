/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           Initial release EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Color variation for a region map chart series
    /// </summary>
    public class ExcelChartExValueColors : XmlHelper
    {
        private ExcelRegionMapChartSerie _series;

        internal ExcelChartExValueColors(ExcelRegionMapChartSerie series, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder) : base (nameSpaceManager, topNode)
        {
            _series = series;
            SchemaNodeOrder = schemaNodeOrder;
        }
        /// <summary>
        /// Number of colors to create the series gradient color scale.
        /// If two colors, the mid color is null.
        /// </summary>
        public eNumberOfColors NumberOfColors 
        { 
            get
            {
                var v=GetXmlNodeString("cx:valueColorPositions/@count");
                if(v=="3")
                {
                    return eNumberOfColors.ThreeColor;
                }
                else
                {
                    return eNumberOfColors.TwoColor;
                }
            }
            set
            {
                SetXmlNodeString("cx:valueColorPositions/@count", ((int)value).ToString(CultureInfo.InvariantCulture));
            }
        }
        ExcelChartExValueColor _minColor = null;
        /// <summary>
        /// The minimum color value.
        /// </summary>
        public ExcelChartExValueColor MinColor 
        {
            get
            {
                if(_minColor==null)
                {
                    _minColor = new ExcelChartExValueColor(NameSpaceManager, TopNode, SchemaNodeOrder, "min");
                }
                return _minColor;
            }
        }
        ExcelChartExValueColor _midColor = null;
        /// <summary>
        /// The mid color value. Null if NumberOfcolors is set to TwoColors
        /// </summary>
        public ExcelChartExValueColor MidColor
        {
            get
            {
                if (NumberOfColors == eNumberOfColors.TwoColor) return null;
                if (_midColor == null)
                {
                    _midColor = new ExcelChartExValueColor(NameSpaceManager, TopNode, SchemaNodeOrder, "mid");
                }
                return _midColor;
            }
        }
        ExcelChartExValueColor _maxColor = null;
        /// <summary>
        /// The maximum color value.
        /// </summary>
        public ExcelChartExValueColor MaxColor
        {
            get
            {
                if (_maxColor == null)
                {
                    _maxColor = new ExcelChartExValueColor(NameSpaceManager, TopNode, SchemaNodeOrder, "max");
                }
                return _maxColor;
            }
        }
    }
}