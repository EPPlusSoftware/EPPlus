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
    public class ExcelChartSerieWithHorizontalErrorBars : ExcelChartSerieWithErrorBars, IDrawingChartErrorBars  
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="chart">Chart series</param>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="node">Topnode</param>
        /// <param name="isPivot">Is pivotchart</param>
        internal ExcelChartSerieWithHorizontalErrorBars(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot) :
            base(chart, ns, node, isPivot)
        {
            foreach(XmlElement errBarsNode in GetNodes("c:errBars"))
            {
                var direction = GetXmlNodeString(errBarsNode, "c:errDir/@val");
                if(direction=="x")
                {
                    ErrorBarsX = new ExcelChartErrorBars(this, errBarsNode);
                }
                else
                {
                    ErrorBars = new ExcelChartErrorBars(this, errBarsNode);
                }
            }

        }
        /// <summary>
        /// Horizontal error bars
        /// <seealso cref="ErrorBarsX"/>
        /// <seealso cref="AddErrorBars(eErrorBarType, eErrorValueType)"/>
        /// </summary>
        public ExcelChartErrorBars ErrorBarsX { get; internal set; }
        /// <summary>
        /// Adds error bars to the chart serie for both vertical and horizontal directions.
        /// </summary>
        /// <param name="barType">The type of error bars</param>
        /// <param name="valueType">The type of value the error bars will show</param>
        public override void AddErrorBars(eErrorBarType barType, eErrorValueType valueType)
        {
            AddErrorBars(barType, valueType, null);
        }
        /// <summary>
        /// Adds error bars to the chart serie for vertical or horizontal directions.
        /// </summary>
        /// <param name="barType">The type of error bars</param>
        /// <param name="valueType">The type of value the error bars will show</param>
        /// <param name="direction">Direction for the error bars. A value of null will add both horizontal and vertical error bars. </param>
        public void AddErrorBars(eErrorBarType barType, eErrorValueType valueType, eErrorBarDirection? direction)
        {
            if (ErrorBars == null && (direction==null || direction==eErrorBarDirection.Y))
            {
                base.AddErrorBars(barType, valueType);
                ErrorBars.SetDirection(eErrorBarDirection.Y);
            }

            if (ErrorBarsX==null && (direction == null || direction == eErrorBarDirection.X))
            {
                ErrorBarsX = GetNewErrorBar(barType, valueType, ErrorBarsX);
                ErrorBarsX.SetDirection(eErrorBarDirection.X);
            }
        }

    }
}
