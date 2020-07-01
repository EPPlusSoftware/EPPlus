/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                   Change
 *************************************************************************************************
  06/23/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Provides easy access to convert the drawing to a it's typed ExcelChart class.
    /// </summary>
    public class ExcelChartAsType
    {
        ExcelDrawing _drawing;
        internal ExcelChartAsType(ExcelDrawing drawing)
        {
            _drawing = drawing;
        }
        /// <summary>
        /// Converts the drawing to it's top level or other nested drawing class.        
        /// </summary>
        /// <typeparam name="T">The type of drawing. T must be inherited from ExcelDrawing</typeparam>
        /// <returns>The drawing as type T</returns>
        public T Type<T>() where T : ExcelChart
        {
            return _drawing as T;
        }

        #region Standard Charts
        /// <summary>
        /// Returns return the drawing as a generic chart. This the base class for all charts.
        /// If this drawing is not a chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a chart</returns>
        public ExcelChart Chart
        {
            get
            {
                return _drawing as ExcelChart;
            }
        }

        /// <summary>
        /// Returns the drawing as an area chart. 
        /// If this drawing is not an area chart, null will be returned
        /// </summary>
        /// <returns>The drawing as an area chart</returns>
        public ExcelAreaChart AreaChart
        {
            get
            {
                return _drawing as ExcelAreaChart;
            }
        }
        /// <summary>
        /// Returns return the drawing as a bar chart. 
        /// If this drawing is not a bar chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a bar chart</returns>
        public ExcelBarChart BarChart
        {
            get
            {
                return _drawing as ExcelBarChart;
            }
        }
        /// <summary>
        /// Returns the drawing as a bubble chart. 
        /// If this drawing is not a bubble chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a bubble chart</returns>
        public ExcelBubbleChart BubbleChart
        {
            get
            {
                return _drawing as ExcelBubbleChart;
            }
        }
        /// <summary>
        /// Returns return the drawing as a doughnut chart. 
        /// If this drawing is not a doughnut chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a doughnut chart</returns>
        public ExcelDoughnutChart DoughnutChart
        {
            get
            {
                return _drawing as ExcelDoughnutChart;
            }
        }
        /// <summary>
        /// Returns the drawing as a PieOfPie or a BarOfPie chart. 
        /// If this drawing is not a PieOfPie or a BarOfPie chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a PieOfPie or a BarOfPie chart</returns>
        public ExcelOfPieChart OfPieChart
        {
            get
            {
                return _drawing as ExcelOfPieChart;
            }
        }
        /// <summary>
        /// Returns the drawing as a pie chart. 
        /// If this drawing is not a pie chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a pie chart</returns>
        public ExcelPieChart PieChart
        {
            get
            {
                return _drawing as ExcelPieChart;
            }
        }
        /// <summary>
        /// Returns return the drawing as a line chart. 
        /// If this drawing is not a line chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a line chart</returns>
        public ExcelLineChart LineChart
        {
            get
            {
                return _drawing as ExcelLineChart;
            }
        }
        /// <summary>
        /// Returns the drawing as a radar chart. 
        /// If this drawing is not a radar chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a radar chart</returns>
        public ExcelRadarChart RadarChart
        {
            get
            {
                return _drawing as ExcelRadarChart;
            }
        }
        /// <summary>
        /// Returns the drawing as a scatter chart. 
        /// If this drawing is not a scatter chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a scatter chart</returns>
        public ExcelScatterChart ScatterChart
        {
            get
            {
                return _drawing as ExcelScatterChart;
            }
        }
        /// <summary>
        /// Returns the drawing as a stock chart. 
        /// If this drawing is not a stock chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a stock chart</returns>
        public ExcelStockChart StockChart
        {
            get
            {
                return _drawing as ExcelStockChart;
            }
        }
        /// <summary>
        /// Returns the drawing as a surface chart. 
        /// If this drawing is not a surface chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a surface chart</returns>
        public ExcelSurfaceChart SurfaceChart
        {
            get
            {
                return _drawing as ExcelSurfaceChart;
            }
        }
        #endregion
        #region ChartEx methods
        /// <summary>
        /// Returns return the drawing as a sunburst chart. 
        /// If this drawing is not a sunburst chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a sunburst chart</returns>
        public ExcelSunburstChart SunburstChart
        {
            get
            {
                return _drawing as ExcelSunburstChart;
            }
        }
        /// <summary>
        /// Returns return the drawing as a treemap chart. 
        /// If this drawing is not a treemap chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a treemap chart</returns>
        public ExcelTreemapChart TreemapChart
        {
            get
            {
                return _drawing as ExcelTreemapChart;
            }
        }
        /// <summary>
        /// Returns return the drawing as a boxwhisker chart. 
        /// If this drawing is not a boxwhisker chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a boxwhisker chart</returns>
        public ExcelBoxWhiskerChart BoxWhiskerChart
        {
            get
            {
                return _drawing as ExcelBoxWhiskerChart;
            }
        }
        /// <summary>
        /// Returns return the drawing as a histogram chart. 
        /// If this drawing is not a histogram chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a histogram Chart</returns>
        public ExcelHistogramChart HistogramChart
        {
            get
            {
                return _drawing as ExcelHistogramChart;
            }
        }
        /// <summary>
        /// Returns return the drawing as a funnel chart. 
        /// If this drawing is not a funnel chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a funnel Chart</returns>
        public ExcelFunnelChart FunnelChart
        {
            get
            {
                return _drawing as ExcelFunnelChart;
            }
        }
        /// <summary>
        /// Returns the drawing as a waterfall chart. 
        /// If this drawing is not a waterfall chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a waterfall chart</returns>
        public ExcelWaterfallChart WaterfallChart
        {
            get
            {
                return _drawing as ExcelWaterfallChart;
            }
        }
        /// <summary>
        /// Returns the drawing as a region map chart. 
        /// If this drawing is not a region map chart, null will be returned
        /// </summary>
        /// <returns>The drawing as a region map chart</returns>
        public ExcelRegionMapChart RegionMapChart
        {
            get
            {
                return _drawing as ExcelRegionMapChart;
            }
        }
        #endregion
    }
}
