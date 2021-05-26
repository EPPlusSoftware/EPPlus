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
using System.Xml;
using OfficeOpenXml.Drawing.Interfaces;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A base class used for chart series that support ErrorBars
    /// </summary>
    public class ExcelChartSerieWithErrorBars : ExcelChartStandardSerie, IDrawingChartErrorBars  
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">Chart series</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        /// <param name="isPivot">Is pivotchart</param>
        internal ExcelChartSerieWithErrorBars(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot) :
            base(chart, ns, node, isPivot)
        {
            var errorNode = GetNode("c:errBars");
            if (errorNode != null) 
            {
                ErrorBars = new ExcelChartErrorBars(this, errorNode);
            }
        }
        /// <summary>
        /// A collection of error bars
        /// <seealso cref="AddErrorBars(eErrorBarType, eErrorValueType)"/>
        /// </summary>
        public ExcelChartErrorBars ErrorBars { get; internal set; } = null;
        /// <summary>
        /// Adds a errorbars to the chart serie
        /// </summary>
        /// <param name="barType"></param>
        /// <param name="valueType"></param>
        public virtual void AddErrorBars(eErrorBarType barType, eErrorValueType valueType)
        {
            ErrorBars = GetNewErrorBar(barType, valueType, ErrorBars);
        }

        internal ExcelChartErrorBars GetNewErrorBar(eErrorBarType barType, eErrorValueType valueType, ExcelChartErrorBars errorBars)
        {
            if (errorBars == null)
            {
                errorBars = new ExcelChartErrorBars(this);
            }
            errorBars.BarType = barType;
            errorBars.ValueType = valueType;
            errorBars.NoEndCap = false;

            _chart.ApplyStyleOnPart(errorBars, _chart.StyleManager?.Style?.ErrorBar);
            return errorBars;
        }

        /// <summary>
        /// Returns true if the serie has Error Bars
        /// </summary>
        /// <returns>True if the serie has Error Bars</returns>
        public bool HasErrorBars()
        {
            return ExistsNode("c:errBars");
        }
    }
}
