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
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Represents a collection of bubble chart series
    /// </summary>
    public sealed class ExcelBubbleChartSeries : ExcelChartSeries<ExcelBubbleChartSerie>
    {
        internal ExcelBubbleChartSeries(ExcelChart chart, XmlNamespaceManager ns, XmlNode node, bool isPivot, List<ExcelChartSerie> list)
        {
            Init(chart, ns, node, isPivot, list);
        }
        /// <summary>
        /// Adds a new serie to a bubble chart
        /// </summary>
        /// <param name="Serie">The Y-Axis range</param>
        /// <param name="XSerie">The X-Axis range</param>
        /// <param name="BubbleSize">The size of the bubbles range. If set to null, a size of 1 is used</param>
        /// <returns></returns>
        public ExcelChartSerie Add(ExcelRangeBase Serie, ExcelRangeBase XSerie, ExcelRangeBase BubbleSize)
        {
            return AddSeries(Serie.FullAddressAbsolute, XSerie.FullAddressAbsolute, BubbleSize?.FullAddressAbsolute);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="SerieAddress">The Y-Axis range</param>
        /// <param name="XSerieAddress">The X-Axis range</param>
        /// <param name="BubbleSizeAddress">The size of the bubbles range. If set to null or String.Empty, a size of 1 is used</param>
        /// <returns></returns>
        public ExcelChartSerie Add(string SerieAddress, string XSerieAddress, string BubbleSizeAddress)
        {
            return AddSeries(SerieAddress, XSerieAddress, BubbleSizeAddress);
        }
    }
}