/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                   Change
 *************************************************************************************************
  06/05/2020         EPPlus Software AB       EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
namespace OfficeOpenXml.Drawing
{
    public partial class ExcelDrawing
    {

        /// <summary>
        /// Returns the drawing as a shape. 
        /// If this drawing is not a shape, null will be returned
        /// </summary>
        /// <returns>The drawing as a shape</returns>
        public ExcelShape AsShape()
        {
            return this as ExcelShape;
        }
        /// <summary>
        /// Returns return the drawing as a picture/image. 
        /// If this drawing is not a picture, null will be returned
        /// </summary>
        /// <returns>The drawing as a picture</returns>
        public ExcelPicture AsPicture()
        {
            return this as ExcelPicture;
        }
        #region Standard Charts
        /// <summary>
        /// Returns the drawing as an area chart. 
        /// If this drawing is not an area chart, null will be returned
        /// </summary>
        /// <returns>The drawing as an area chart</returns>
        public ExcelAreaChart AsAreaChart()
        {
            return this as ExcelAreaChart;
        }
        /// <summary>
        /// Returns return the drawing as a bar chart. 
        /// If this drawing is not a bar chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a bar chart</returns>
        public ExcelBarChart AsBarChart()
        {
            return this as ExcelBarChart;
        }
        /// <summary>
        /// Returns the drawing as a bubble chart. 
        /// If this drawing is not a bubble chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a bubble chart</returns>
        public ExcelBubbleChart AsBubbleChart()
        {
            return this as ExcelBubbleChart;
        }
        /// <summary>
        /// Returns return the drawing as a doughnut chart. 
        /// If this drawing is not a doughnut chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a doughnut chart</returns>
        public ExcelDoughnutChart AsDoughnutChart()
        {
            return this as ExcelDoughnutChart;
        }
        /// <summary>
        /// Returns the drawing as a PieOfPie or a BarOfPie chart. 
        /// If this drawing is not a PieOfPie or a BarOfPie chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a PieOfPie or a BarOfPie chart</returns>
        public ExcelOfPieChart AsOfPieChart()
        {
            return this as ExcelOfPieChart;
        }
        /// <summary>
        /// Returns the drawing as a pie chart. 
        /// If this drawing is not a pie chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a pie chart</returns>
        public ExcelPieChart AsPieChart()
        {
            return this as ExcelPieChart;
        }
        /// <summary>
        /// Returns return the drawing as a line chart. 
        /// If this drawing is not a line chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a line chart</returns>
        public ExcelLineChart AsLineChart()
        {
            return this as ExcelLineChart;
        }
        /// <summary>
        /// Returns the drawing as a radar chart. 
        /// If this drawing is not a radar chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a radar chart</returns>
        public ExcelRadarChart AsRadarChart()
        {
            return this as ExcelRadarChart;
        }
        /// <summary>
        /// Returns the drawing as a scatter chart. 
        /// If this drawing is not a scatter chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a scatter chart</returns>
        public ExcelScatterChart AsScatterChart()
        {
            return this as ExcelScatterChart;
        }
        /// <summary>
        /// Returns the drawing as a stock chart. 
        /// If this drawing is not a stock chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a stock chart</returns>
        public ExcelStockChart AsStockChart()
        {
            return this as ExcelStockChart;
        }
        /// <summary>
        /// Returns the drawing as a surface chart. 
        /// If this drawing is not a surface chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a surface chart</returns>
        public ExcelSurfaceChart AsSurfaceChart()
        {
            return this as ExcelSurfaceChart;
        }
        #endregion
        #region ChartEx methods
        /// <summary>
        /// Returns return the drawing as a sunburst chart. 
        /// If this drawing is not a sunburst chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a sunburst chart</returns>
        public ExcelSunburstChart AsSunburstChart()
        {
            return this as ExcelSunburstChart;
        }
        /// <summary>
        /// Returns return the drawing as a treemap chart. 
        /// If this drawing is not a treemap chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a treemap chart</returns>
        public ExcelTreemapChart AsTreemapChart()
        {
            return this as ExcelTreemapChart;
        }
        /// <summary>
        /// Returns return the drawing as a boxwhisker chart. 
        /// If this drawing is not a boxwhisker chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a boxwhisker chart</returns>
        public ExcelBoxWhiskerChart AsBoxWhiskerChart()
        {
            return this as ExcelBoxWhiskerChart;
        }
        /// <summary>
        /// Returns return the drawing as a histogram chart. 
        /// If this drawing is not a histogram chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a histogram Chart</returns>
        public ExcelHistogramChart AsHistogramChart()
        {
            return this as ExcelHistogramChart;
        }
        /// <summary>
        /// Returns return the drawing as a funnel chart. 
        /// If this drawing is not a funnel chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a funnel Chart</returns>
        public ExcelFunnelChart AsFunnelChart()
        {
            return this as ExcelFunnelChart;
        }
        /// <summary>
        /// Returns the drawing as a waterfall chart. 
        /// If this drawing is not a waterfall chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a waterfall chart</returns>
        public ExcelWaterfallChart AsWaterfallChart()
        {
            return this as ExcelWaterfallChart;
        }
        /// <summary>
        /// Returns the drawing as a region map chart. 
        /// If this drawing is not a region map chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a region map chart</returns>
        public ExcelRegionMapChart AsRegionMapChart()
        {
            return this as ExcelRegionMapChart;
        }
        #endregion
    }
}
