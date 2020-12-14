/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB           EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Utils.Extensions;
using System.Xml;
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A series for an Waterfall Chart
    /// </summary>
    public class ExcelWaterfallChartSerie : ExcelChartExSerie
    {
        internal ExcelWaterfallChartSerie(ExcelChartEx chart, XmlNamespaceManager ns, XmlNode node) : base(chart, ns, node)
        {

        }

        const string _connectorLinesPath = "cx:layoutPr/cx:visibility/@connectorLines";
        /// <summary>
        /// The visibility of connector lines between data points
        /// </summary>
        public bool ShowConnectorLines
        {
            get
            {
                return GetXmlNodeBool($"{_connectorLinesPath}");
            }
            set
            {
                SetXmlNodeBool($"{_connectorLinesPath}", value);
            }
        }
    }
}
