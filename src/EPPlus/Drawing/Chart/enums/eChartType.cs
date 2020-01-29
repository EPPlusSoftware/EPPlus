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
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Chart type
    /// </summary>
    public enum eChartType
    {
        /// <summary>
        /// An 3D area chart
        /// </summary>
        Area3D = -4098,
        /// <summary>
        /// A stacked area 3D chart
        /// </summary>
        AreaStacked3D = 78,
        /// <summary>
        /// A 100% stacked 3D area chart
        /// </summary>
        AreaStacked1003D = 79,
        /// <summary>
        /// A clustered 3D bar chart
        /// </summary>
        BarClustered3D = 60,
        /// <summary>
        /// A stacked 3D bar chart
        /// </summary>
        BarStacked3D = 61,
        /// <summary>
        /// A 100% stacked 3D bar chart
        /// </summary>
        BarStacked1003D = 62,
        /// <summary>
        /// A 3D column chart
        /// </summary>
        Column3D = -4100,
        /// <summary>
        /// A clustered 3D column chart
        /// </summary>
        ColumnClustered3D = 54,
        /// <summary>
        /// A stacked 3D column chart
        /// </summary>
        ColumnStacked3D = 55,
        /// <summary>
        /// A 100% stacked 3D column chart
        /// </summary>
        ColumnStacked1003D = 56,
        /// <summary>
        /// A 3D line chart
        /// </summary>
        Line3D = -4101,
        /// <summary>
        /// A 3D pie chart
        /// </summary>
        Pie3D = -4102,
        /// <summary>
        /// A exploded 3D pie chart
        /// </summary>
        PieExploded3D = 70,
        /// <summary>
        /// An area chart
        /// </summary>
        Area = 1,
        /// <summary>
        /// A stacked area chart
        /// </summary>
        AreaStacked = 76,
        /// <summary>
        /// A 100% stacked area chart
        /// </summary>
        AreaStacked100 = 77,
        /// <summary>
        /// A clustered bar chart
        /// </summary>
        BarClustered = 57,
        /// <summary>
        /// A bar of pie chart
        /// </summary>
        BarOfPie = 71,
        /// <summary>
        /// A stacked bar chart
        /// </summary>
        BarStacked = 58,
        /// <summary>
        /// A 100% stacked bar chart
        /// </summary>
        BarStacked100 = 59,
        /// <summary>
        /// A bubble chart 
        /// </summary>
        Bubble = 15,
        /// <summary>
        /// A 3D bubble chart 
        /// </summary>
        Bubble3DEffect = 87,
        /// <summary>
        /// A clustered column chart 
        /// </summary>
        ColumnClustered = 51,
        /// <summary>
        /// A stacked column chart
        /// </summary>
        ColumnStacked = 52,
        /// <summary>
        /// A 100% stacked column chart
        /// </summary>
        ColumnStacked100 = 53,
        /// <summary>
        /// A clustered cone bar chart
        /// </summary>
        ConeBarClustered = 102,
        /// <summary>
        /// A stacked cone bar chart
        /// </summary>
        ConeBarStacked = 103,
        /// <summary>
        /// A 100% stacked cone bar chart
        /// </summary>
        ConeBarStacked100 = 104,
        /// <summary>
        /// A cone column chart 
        /// </summary>
        ConeCol = 105,
        /// <summary>
        /// A clustered cone column chart 
        /// </summary>
        ConeColClustered = 99,
        /// <summary>
        /// A stacked cone column chart 
        /// </summary>
        ConeColStacked = 100,
        /// <summary>
        /// A 100% stacked cone column chart 
        /// </summary>
        ConeColStacked100 = 101,
        /// <summary>
        /// A clustered cylinder bar chart
        /// </summary>
        CylinderBarClustered = 95,
        /// <summary>
        /// A stacked cylinder bar chart
        /// </summary>
        CylinderBarStacked = 96,
        /// <summary>
        /// A 100% stacked cylinder bar chart
        /// </summary>
        CylinderBarStacked100 = 97,
        /// <summary>
        /// A cylinder column chart
        /// </summary>
        CylinderCol = 98,
        /// <summary>
        /// A clustered cylinder column chart
        /// </summary>
        CylinderColClustered = 92,
        /// <summary>
        /// A stacked cylinder column chart
        /// </summary>
        CylinderColStacked = 93,
        /// <summary>
        /// A 100% stacked cylinder column chart
        /// </summary>
        CylinderColStacked100 = 94,
        /// <summary>
        /// A doughnut chart
        /// </summary>
        Doughnut = -4120,
        /// <summary>
        /// An exploded doughnut chart
        /// </summary>
        DoughnutExploded = 80,
        /// <summary>
        /// A line chart
        /// </summary>
        Line = 4,
        /// <summary>
        /// A line chart with markers
        /// </summary>
        LineMarkers = 65,
        /// <summary>
        /// A stacked line chart with markers
        /// </summary>
        LineMarkersStacked = 66,
        /// <summary>
        /// A 100% stacked line chart with markers
        /// </summary>
        LineMarkersStacked100 = 67,
        /// <summary>
        /// A stacked line chart
        /// </summary>
        LineStacked = 63,
        /// <summary>
        /// A 100% stacked line chart
        /// </summary>
        LineStacked100 = 64,
        /// <summary>
        /// A pie chart
        /// </summary>
        Pie = 5,
        /// <summary>
        /// An exploded pie chart
        /// </summary>
        PieExploded = 69,
        /// <summary>
        /// A pie of pie chart
        /// </summary>
        PieOfPie = 68,
        /// <summary>
        /// A clustered pyramid bar chart
        /// </summary>
        PyramidBarClustered = 109,
        /// <summary>
        /// A stacked pyramid bar chart
        /// </summary>
        PyramidBarStacked = 110,
        /// <summary>
        /// A 100% stacked pyramid bar chart
        /// </summary>
        PyramidBarStacked100 = 111,
        /// <summary>
        /// A stacked pyramid column chart
        /// </summary>
        PyramidCol = 112,
        /// <summary>
        /// A clustered pyramid column chart
        /// </summary>
        PyramidColClustered = 106,
        /// <summary>
        /// A stacked pyramid column chart
        /// </summary>
        PyramidColStacked = 107,
        /// <summary>
        /// A 100% stacked pyramid column chart
        /// </summary>
        PyramidColStacked100 = 108,
        /// <summary>
        /// A radar chart
        /// </summary>
        Radar = -4151,
        /// <summary>
        /// A filled radar chart
        /// </summary>
        RadarFilled = 82,
        /// <summary>
        /// A radar chart with markers
        /// </summary>
        RadarMarkers = 81,
        /// <summary>
        /// Stock chart, not supported in EPPlus
        /// </summary>
        StockHLC =88,
        /// <summary>
        /// Stock chart, not supported in EPPlus
        /// </summary>
        StockOHLC = 89,
        /// <summary>
        /// Stock chart, not supported in EPPlus
        /// </summary>
        StockVHLC = 90,
        /// <summary>
        /// Stock chart, not supported in EPPlus
        /// </summary>
        StockVOHLC = 91,
        /// <summary>
        /// A surface chart
        /// </summary>
        Surface = 83,
        /// <summary>
        /// A surface chart, top view
        /// </summary>
        SurfaceTopView = 85,
        /// <summary>
        /// A surface chart, top view and wireframe
        /// </summary>
        SurfaceTopViewWireframe = 86,
        /// <summary>
        /// A surface chart, wireframe
        /// </summary>
        SurfaceWireframe = 84,
        /// <summary>
        /// A XY scatter chart
        /// </summary>
        XYScatter = -4169,
        /// <summary>
        /// A scatter line chart with markers
        /// </summary>
        XYScatterLines = 74,
        /// <summary>
        /// A scatter line chart with no markers
        /// </summary>
        XYScatterLinesNoMarkers = 75,
        /// <summary>
        /// A scatter line chart with markers and smooth lines
        /// </summary>
        XYScatterSmooth = 72,
        /// <summary>
        /// A scatter line chart with no markers and smooth lines
        /// </summary>
        XYScatterSmoothNoMarkers = 73
    }
}