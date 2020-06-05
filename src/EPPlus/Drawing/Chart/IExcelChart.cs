using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.ThreeD;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    public interface IExcelChart 
    {
        /// <summary>
        /// Manage style settings for the chart
        /// </summary>
        ExcelChartStyleManager StyleManager { get; }
        /// <summary>
        /// Reference to the worksheet
        /// </summary>
        ExcelWorksheet WorkSheet { get; }
        /// <summary>
        /// The chart xml document
        /// </summary>
        XmlDocument ChartXml { get; }
        /// <summary>
        /// Type of chart
        /// </summary>
        eChartType ChartType { get; }
        /// <summary>
        /// Titel of the chart
        /// </summary>
        ExcelChartTitle Title { get; }
        /// <summary>
        /// True if the chart has a title
        /// </summary>
        bool HasTitle { get; }
        /// <summary>
        /// If the chart has a legend
        /// </summary>
        bool HasLegend { get; }
        /// <summary>
        /// Remove the title from the chart
        /// </summary>
        void DeleteTitle();
        /// <summary>
        /// Chart series
        /// </summary>
        ExcelChartSeries<ExcelChartSerie> Series { get; }
        /// <summary>
        /// An array containg all axis of all Charttypes
        /// </summary>
        ExcelChartAxis[] Axis { get; }
        /// <summary>
        /// The X Axis
        /// </summary>
        ExcelChartAxis XAxis { get; }
        /// <summary>
        /// The Y Axis
        /// </summary>
        ExcelChartAxis YAxis { get; }
        /// <summary>
        /// If true the charttype will use the secondary axis.
        /// The chart must contain a least one other charttype that uses the primary axis.
        /// </summary>
        bool UseSecondaryAxis { get; }
        /// <summary>
        /// The build-in chart styles. 
        /// </summary>
        eChartStyle Style { get; }
        /// <summary>
        /// Border rounded corners
        /// </summary>
        bool RoundedCorners { get; set; }
        /// <summary>
        /// Show data in hidden rows and columns
        /// </summary>
        bool ShowHiddenData { get; set; }
        /// <summary>
        /// Specifies the possible ways to display blanks
        /// </summary>
        eDisplayBlanksAs DisplayBlanksAs { get; set; }
        /// <summary>
        /// Specifies data labels over the maximum of the chart shall be shown
        /// </summary>
        bool ShowDataLabelsOverMaximum { get; set; }
        /// <summary>
        /// Plotarea
        /// </summary>
        ExcelChartPlotArea PlotArea { get; }
        /// <summary>
        /// Legend
        /// </summary>
        ExcelChartLegend Legend { get; }
        /// <summary>
        /// Border
        /// </summary>
        ExcelDrawingBorder Border { get; }
        /// <summary>
        /// 3D properties
        /// </summary>
        ExcelDrawing3D ThreeD { get; }
        /// <summary>
        /// Access to font properties
        /// </summary>
        ExcelTextFont Font { get; }
        /// <summary>
        /// Access to text body properties
        /// </summary>
        ExcelTextBody TextBody { get; }
        /// <summary>
        /// 3D-settings
        /// </summary>
        ExcelView3D View3D { get; }
        /// <summary>
        /// Specifies the kind of grouping for a column, line, or area chart
        /// </summary>
        eGrouping Grouping { get; }
        /// <summary>
        /// If the chart has only one serie this varies the colors for each point.
        /// </summary>
        bool VaryColors { get; }
        /// <summary>
        /// If the chart is a pivochart this is the pivotable used as source.
        /// </summary>
        ExcelPivotTable PivotTableSource { get; }
    }
}