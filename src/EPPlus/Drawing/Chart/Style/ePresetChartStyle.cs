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
namespace OfficeOpenXml.Drawing.Chart.Style
{
    /// <summary>
    /// Maps to Excel's built-in chart styles, primary for charts with one data serie. 
    /// Note that Excel changes chart type depending on many parameters, like number of series, axis type and more, so it will not always match the number in Excel.
    /// To be certain of getting the correct style use the chart style number of the style you want to apply
    /// For charts with more than one series use <see cref="ePresetChartStyleMultiSeries"/>
    /// By default the styles are loaded into the StyleLibrary.You can also load your own with your own id's.
    /// Styles are fetched from the StyleLibrary by the id provided in this enum. 
    /// <seealso cref="ExcelChartStyleManager.StyleLibrary" />
    /// </summary>
    public enum ePresetChartStyle
    {
        /// <summary>
        /// 3D Area Chart style 1
        /// </summary>
        Area3dChartStyle1 = 276,
        /// <summary>
        /// 3D Area Chart style 2
        /// </summary>
        Area3dChartStyle2 = 312,
        /// <summary>
        /// 3D Area Chart style 3
        /// </summary>
        Area3dChartStyle3 = 313,
        /// <summary>
        /// 3D Area Chart style 4
        /// </summary>
        Area3dChartStyle4 = 314,
        /// <summary>
        /// 3D Area Chart style 5
        /// </summary>
        Area3dChartStyle5 = 280,
        /// <summary>
        /// 3D Area Chart style 6
        /// </summary>
        Area3dChartStyle6 = 281,
        /// <summary>
        /// 3D Area Chart style 7
        /// </summary>
        Area3dChartStyle7 = 282,
        /// <summary>
        /// 3D Area Chart style 8
        /// </summary>
        Area3dChartStyle8 = 315,
        /// <summary>
        /// 3D Area Chart style 9
        /// </summary>
        Area3dChartStyle9 = 316,
        /// <summary>
        /// 3D Area Chart style 10
        /// </summary>
        Area3dChartStyle10 = 350,
        /// <summary>
        /// Area Chart style 1
        /// </summary>
        AreaChartStyle1 = 276,
        /// <summary>
        /// Area Chart style 2
        /// </summary>
        AreaChartStyle2 = 277,
        /// <summary>
        /// Area Chart style 3
        /// </summary>
        AreaChartStyle3 = 278,
        /// <summary>
        /// Area Chart style 4
        /// </summary>
        AreaChartStyle4 = 279,
        /// <summary>
        /// Area Chart style 5
        /// </summary>
        AreaChartStyle5 = 280,
        /// <summary>
        /// Area Chart style 6
        /// </summary>
        AreaChartStyle6 = 281,
        /// <summary>
        /// Area Chart style 7
        /// </summary>
        AreaChartStyle7 = 282,
        /// <summary>
        /// Area Chart style 8
        /// </summary>
        AreaChartStyle8 = 283,
        /// <summary>
        /// Area Chart style 9
        /// </summary>
        AreaChartStyle9 = 284,
        /// <summary>
        /// Area Chart style 10
        /// </summary>
        AreaChartStyle10 = 285,
        /// <summary>
        /// Area Chart style 11
        /// </summary>
        AreaChartStyle11 = 346,
        /// <summary>
        /// Bar 3d Chart Style 1
        /// </summary>
        Bar3dChartStyle1 = 286,
        /// <summary>
        /// Bar 3d Chart Style 2
        /// </summary>
        Bar3dChartStyle2 = 287,
        /// <summary>
        /// Bar 3d Chart Style 3
        /// </summary>
        Bar3dChartStyle3 = 288,
        /// <summary>
        /// Bar 3d Chart Style 4
        /// </summary>
        Bar3dChartStyle4 = 289,
        /// <summary>
        /// Bar 3d Chart Style 5
        /// </summary>
        Bar3dChartStyle5 = 290,
        /// <summary>
        /// Bar 3d Chart Style 6
        /// </summary>
        Bar3dChartStyle6 = 291,
        /// <summary>
        /// Bar 3d Chart Style 7
        /// </summary>
        Bar3dChartStyle7 = 292,
        /// <summary>
        /// Bar 3d Chart Style 8
        /// </summary>
        Bar3dChartStyle8 = 349,
        /// <summary>
        /// Bar 3d Chart Style 9
        /// </summary>
        Bar3dChartStyle9 = 294,
        /// <summary>
        /// Bar 3d Chart Style 10
        /// </summary>
        Bar3dChartStyle10 = 295,
        /// <summary>
        /// Bar 3d Chart Style 11
        /// </summary>
        Bar3dChartStyle11 = 296,
        /// <summary>
        /// Bar 3d Chart Style 12
        /// </summary>
        Bar3dChartStyle12 = 347,
        /// <summary>
        /// Bar Chart style 1
        /// </summary>
        BarChartStyle1 = 216,
        /// <summary>
        /// Bar Chart style 2
        /// </summary>
        BarChartStyle2 = 217,
        /// <summary>
        /// Bar Chart style 3
        /// </summary>
        BarChartStyle3 = 218,
        /// <summary>
        /// Bar Chart style 4
        /// </summary>
        BarChartStyle4 = 219,
        /// <summary>
        /// Bar Chart style 5
        /// </summary>
        BarChartStyle5 = 220,
        /// <summary>
        /// Bar Chart style 6
        /// </summary>
        BarChartStyle6 = 221,
        /// <summary>
        /// Bar Chart style 7
        /// </summary>
        BarChartStyle7 = 222,
        /// <summary>
        /// Bar Chart style 8
        /// </summary>
        BarChartStyle8 = 223,
        /// <summary>
        /// Bar Chart style 9
        /// </summary>
        BarChartStyle9 = 224,
        /// <summary>
        /// Bar Chart style 10
        /// </summary>
        BarChartStyle10 = 225,
        /// <summary>
        /// Bar Chart style 11
        /// </summary>
        BarChartStyle11 = 339,
        /// <summary>
        /// Bar Chart style 12
        /// </summary>
        BarChartStyle12 = 226,
        /// <summary>
        /// Bar Chart style 13
        /// </summary>
        BarChartStyle13 = 341,
        /// <summary>
        /// Bubble Chart Style 1
        /// </summary>
        BubbleChartStyle1 = 269,
        /// <summary>
        /// Bubble Chart Style 2
        /// </summary>
        BubbleChartStyle2 = 270,
        /// <summary>
        /// Bubble 3d Chart Style 1
        /// </summary>
        Bubble3dChartStyle1 = 269,
        /// <summary>
        /// Bubble 3d Chart Style 2
        /// </summary>
        Bubble3dChartStyle2 = 270,
        /// <summary>
        /// Bubble 3d Chart Style 3
        /// </summary>
        Bubble3dChartStyle3 = 272,
        /// <summary>
        /// Bubble 3d Chart Style 4
        /// </summary>
        Bubble3dChartStyle4 = 246,
        /// <summary>
        /// Bubble 3d Chart Style 5
        /// </summary>
        Bubble3dChartStyle5 = 242,
        /// <summary>
        /// Bubble 3d Chart Style 6
        /// </summary>
        Bubble3dChartStyle6 = 273,
        /// <summary>
        /// Bubble 3d Chart Style 7
        /// </summary>
        Bubble3dChartStyle7 = 248,
        /// <summary>
        /// Bubble 3d Chart Style 8
        /// </summary>
        Bubble3dChartStyle8 = 275,
        /// <summary>
        /// Bubble 3d Chart Style 8
        /// </summary>
        Bubble3dChartStyle9 = 343,
        /// <summary>
        /// Bubble Chart Style 3
        /// </summary>
        BubbleChartStyle3 = 271,
        /// <summary>
        /// Bubble Chart Style 4
        /// </summary>
        BubbleChartStyle4 = 272,
        /// <summary>
        /// Bubble Chart Style 5
        /// </summary>
        BubbleChartStyle5 = 246,
        /// <summary>
        /// Bubble Chart Style 6
        /// </summary>
        BubbleChartStyle6 = 242,
        /// <summary>
        /// Bubble Chart Style 7
        /// </summary>
        BubbleChartStyle7 = 273,
        /// <summary>
        /// Bubble Chart Style 8
        /// </summary>
        BubbleChartStyle8 = 248,
        /// <summary>
        /// Bubble Chart Style 9
        /// </summary>
        BubbleChartStyle9 = 274,
        /// <summary>
        /// Bubble Chart Style 10
        /// </summary>
        BubbleChartStyle10 = 275,
        /// <summary>
        /// Bubble Chart Style 11
        /// </summary>
        BubbleChartStyle11 = 343,
        /// <summary>
        /// Column 3d Chart Style 1
        /// </summary>
        Column3dChartStyle1 = 286,
        /// <summary>
        /// Column 3d Chart Style 2
        /// </summary>
        Column3dChartStyle2 = 287,
        /// <summary>
        /// Column 3d Chart Style 3
        /// </summary>
        Column3dChartStyle3 = 288,
        /// <summary>
        /// Column 3d Chart Style 4
        /// </summary>
        Column3dChartStyle4 = 289,
        /// <summary>
        /// Column 3d Chart Style 5
        /// </summary>
        Column3dChartStyle5 = 290,
        /// <summary>
        /// Column 3d Chart Style 6
        /// </summary>
        Column3dChartStyle6 = 291,
        /// <summary>
        /// Column 3d Chart Style 7
        /// </summary>
        Column3dChartStyle7 = 292,
        /// <summary>
        /// Column 3d Chart Style 8
        /// </summary>
        Column3dChartStyle8 = 293,
        /// <summary>
        /// Column 3d Chart Style 9
        /// </summary>
        Column3dChartStyle9 = 294,
        /// <summary>
        /// Column 3d Chart Style 10
        /// </summary>
        Column3dChartStyle10 = 295,
        /// <summary>
        /// Column 3d Chart Style 11
        /// </summary>
        Column3dChartStyle11 = 296,
        /// <summary>
        /// Column 3d Chart Style 12
        /// </summary>
        Column3dChartStyle12 = 347,
        /// <summary>
        /// Column Chart style 1
        /// </summary>
        ColumnChartStyle1 = 201,
        /// <summary>
        /// Column Chart style 2
        /// </summary>
        ColumnChartStyle2 = 202,
        /// <summary>
        /// Column Chart style 3
        /// </summary>
        ColumnChartStyle3 = 203,
        /// <summary>
        /// Column Chart style 4
        /// </summary>
        ColumnChartStyle4 = 204,
        /// <summary>
        /// Column Chart style 5
        /// </summary>
        ColumnChartStyle5 = 205,
        /// <summary>
        /// Column Chart style 6
        /// </summary>
        ColumnChartStyle6 = 206,
        /// <summary>
        /// Column Chart style 7
        /// </summary>
        ColumnChartStyle7 = 207,
        /// <summary>
        /// Column Chart style 8
        /// </summary>
        ColumnChartStyle8 = 208,
        /// <summary>
        /// Column Chart style 9
        /// </summary>
        ColumnChartStyle9 = 209,
        /// <summary>
        /// Column Chart style 10
        /// </summary>
        ColumnChartStyle10 = 210,
        /// <summary>
        /// Column Chart style 11
        /// </summary>
        ColumnChartStyle11 = 211,
        /// <summary>
        /// Column Chart style 12
        /// </summary>
        ColumnChartStyle12 = 212,
        /// <summary>
        /// Column Chart style 13
        /// </summary>
        ColumnChartStyle13 = 213,
        /// <summary>
        /// Column Chart style 14
        /// </summary>
        ColumnChartStyle14 = 214,
        /// <summary>
        /// Column Chart style 15
        /// </summary>
        ColumnChartStyle15 = 215,
        /// <summary>
        /// Column Chart style 16
        /// </summary>
        ColumnChartStyle16 = 340,
        /// <summary>
        /// Custom Combined Chart Style 1
        /// </summary>
        ComboChartStyle1 = 322,
        /// <summary>
        /// Custom Combined Chart Style 2
        /// </summary>
        ComboChartStyle2 = 323,
        /// <summary>
        /// Custom Combined Chart Style 3
        /// </summary>
        ComboChartStyle3 = 325,
        /// <summary>
        /// Custom Combined Chart Style 4
        /// </summary>
        ComboChartStyle4 = 326,
        /// <summary>
        /// Custom Combined Chart Style 5
        /// </summary>
        ComboChartStyle5 = 221,
        /// <summary>
        /// Custom Combined Chart Style 6
        /// </summary>
        ComboChartStyle6 = 328,
        /// <summary>
        /// Custom Combined Chart Style 7
        /// </summary>
        ComboChartStyle7 = 225,
        /// <summary>
        /// Custom Combined Chart Style 8
        /// </summary>
        ComboChartStyle8 = 352,
        /// <summary>
        /// Doughnut Chart Style 1
        /// </summary>
        DoughnutChartStyle1 = 251,
        /// <summary>
        /// Doughnut Chart Style 2
        /// </summary>
        DoughnutChartStyle2 = 252,
        /// <summary>
        /// Doughnut Chart Style 3
        /// </summary>
        DoughnutChartStyle3 = 253,
        /// <summary>
        /// Doughnut Chart Style 4
        /// </summary>
        DoughnutChartStyle4 = 254,
        /// <summary>
        /// Doughnut Chart Style 5
        /// </summary>
        DoughnutChartStyle5 = 255,
        /// <summary>
        /// Doughnut Chart Style 6
        /// </summary>
        DoughnutChartStyle6 = 256,
        /// <summary>
        /// Doughnut Chart Style 7
        /// </summary>
        DoughnutChartStyle7 = 257,
        /// <summary>
        /// Doughnut Chart Style 8
        /// </summary>
        DoughnutChartStyle8 = 258,
        /// <summary>
        /// Doughnut Chart Style 9
        /// </summary>
        DoughnutChartStyle9 = 260,
        /// <summary>
        /// Doughnut Chart Style 10
        /// </summary>
        DoughnutChartStyle10 = 261,
        /// <summary>
        /// Line 3d Chart style 1
        /// </summary>
        Line3dChartStyle1 = 307,
        /// <summary>
        /// Line 3d Chart style 2
        /// </summary>
        Line3dChartStyle2 = 311,
        /// <summary>
        /// Line 3d Chart style 3
        /// </summary>
        Line3dChartStyle3 = 308,
        /// <summary>
        /// Line 3d Chart style 4
        /// </summary>
        Line3dChartStyle4 = 309,
        /// <summary>
        /// Line Chart style 1
        /// </summary>
        LineChartStyle1 = 227,
        /// <summary>
        /// Line Chart style 2
        /// </summary>
        LineChartStyle2 = 228,
        /// <summary>
        /// Line Chart style 3
        /// </summary>
        LineChartStyle3 = 229,
        /// <summary>
        /// Line Chart style 4
        /// </summary>
        LineChartStyle4 = 230,
        /// <summary>
        /// Line Chart style 5
        /// </summary>
        LineChartStyle5 = 231,
        /// <summary>
        /// Line Chart style 6
        /// </summary>
        LineChartStyle6 = 232,
        /// <summary>
        /// Line Chart style 7
        /// </summary>
        LineChartStyle7 = 233,
        /// <summary>
        /// Line Chart style 8
        /// </summary>
        LineChartStyle8 = 234,
        /// <summary>
        /// Line Chart style 9
        /// </summary>
        LineChartStyle9 = 235,
        /// <summary>
        /// Line Chart style 10
        /// </summary>
        LineChartStyle10 = 236,
        /// <summary>
        /// Line Chart style 11
        /// </summary>
        LineChartStyle11 = 237,
        /// <summary>
        /// Line Chart style 12
        /// </summary>
        LineChartStyle12 = 238,
        /// <summary>
        /// Line Chart style 13
        /// </summary>
        LineChartStyle13 = 239,
        /// <summary>
        /// Line Chart style 14
        /// </summary>
        LineChartStyle14 = 332,
        /// <summary>
        /// Line Chart style 15
        /// </summary>
        LineChartStyle15 = 342,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 1
        /// </summary>
        OfPieChartStyle1 = 333,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 2
        /// </summary>
        OfPieChartStyle2 = 252,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 3
        /// </summary>
        OfPieChartStyle3 = 334,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 4
        /// </summary>
        OfPieChartStyle4 = 335,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 5
        /// </summary>
        OfPieChartStyle5 = 336,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 6
        /// </summary>
        OfPieChartStyle6 = 337,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 7
        /// </summary>
        OfPieChartStyle7 = 338,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 8
        /// </summary>
        OfPieChartStyle8 = 258,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 9
        /// </summary>
        OfPieChartStyle9 = 259,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 10
        /// </summary>
        OfPieChartStyle10 = 260,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 11
        /// </summary>
        OfPieChartStyle11 = 261,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 12
        /// </summary>
        OfPieChartStyle12 = 344,
        /// <summary>
        /// Pie 3d Chart Style 1
        /// </summary>
        Pie3dChartStyle1 = 262,
        /// <summary>
        /// Pie 3d Chart Style 2
        /// </summary>
        Pie3dChartStyle2 = 263,
        /// <summary>
        /// Pie 3d Chart Style 3
        /// </summary>
        Pie3dChartStyle3 = 264,
        /// <summary>
        /// Pie 3d Chart Style 4
        /// </summary>
        Pie3dChartStyle4 = 265,
        /// <summary>
        /// Pie 3d Chart Style 5
        /// </summary>
        Pie3dChartStyle5 = 266,
        /// <summary>
        /// Pie 3d Chart Style 6
        /// </summary>
        Pie3dChartStyle6 = 267,
        /// <summary>
        /// Pie 3d Chart Style 7
        /// </summary>
        Pie3dChartStyle7 = 268,
        /// <summary>
        /// Pie 3d Chart Style 8
        /// </summary>
        Pie3dChartStyle8 = 259,
        /// <summary>
        /// Pie 3d Chart Style 9
        /// </summary>
        Pie3dChartStyle9 = 261,
        /// <summary>
        /// Pie 3d Chart Style 10
        /// </summary>
        Pie3dChartStyle10 = 345,
        /// <summary>
        /// Pie Chart Style 1
        /// </summary>
        PieChartStyle1 = 251,
        /// <summary>
        /// Pie Chart Style 2
        /// </summary>
        PieChartStyle2 = 252,
        /// <summary>
        /// Pie Chart Style 3
        /// </summary>
        PieChartStyle3 = 253,
        /// <summary>
        /// Pie Chart Style 4
        /// </summary>
        PieChartStyle4 = 254,
        /// <summary>
        /// Pie Chart Style 5
        /// </summary>
        PieChartStyle5 = 255,
        /// <summary>
        /// Pie Chart Style 6
        /// </summary>
        PieChartStyle6 = 256,
        /// <summary>
        /// Pie Chart Style 7
        /// </summary>
        PieChartStyle7 = 257,
        /// <summary>
        /// Pie Chart Style 8
        /// </summary>
        PieChartStyle8 = 258,
        /// <summary>
        /// Pie Chart Style 9
        /// </summary>
        PieChartStyle9 = 259,
        /// <summary>
        /// Pie Chart Style 10
        /// </summary>
        PieChartStyle10 = 260,
        /// <summary>
        /// Pie Chart style 11
        /// </summary>
        PieChartStyle11 = 261,
        /// <summary>
        /// Pie Chart style 12
        /// </summary>
        PieChartStyle12 = 344,
        /// <summary>
        /// Radar Chart style 1
        /// </summary>
        RadarChartStyle1 = 317,
        /// <summary>
        /// Radar Chart style 2
        /// </summary>
        RadarChartStyle2 = 318,
        /// <summary>
        /// Radar Chart style 3
        /// </summary>
        RadarChartStyle3 = 206,
        /// <summary>
        /// Radar Chart style 4
        /// </summary>
        RadarChartStyle4 = 319,
        /// <summary>
        /// Radar Chart style 5
        /// </summary>
        RadarChartStyle5 = 320,
        /// <summary>
        /// Radar Chart style 6
        /// </summary>
        RadarChartStyle6 = 207,
        /// <summary>
        /// Radar Chart style 7
        /// </summary>
        RadarChartStyle7 = 321,
        /// <summary>
        /// Radar Chart style 8
        /// </summary>
        RadarChartStyle8 = 351,
        /// <summary>
        /// Scatter Chart style 1
        /// </summary>
        ScatterChartStyle1 = 240,
        /// <summary>
        /// Scatter Chart style 2
        /// </summary>
        ScatterChartStyle2 = 241,
        /// <summary>
        /// Scatter Chart style 3
        /// </summary>
        ScatterChartStyle3 = 242,
        /// <summary>
        /// Scatter Chart style 4
        /// </summary>
        ScatterChartStyle4 = 243,
        /// <summary>
        /// Scatter Chart style 5
        /// </summary>
        ScatterChartStyle5 = 244,
        /// <summary>
        /// Scatter Chart style 6
        /// </summary>
        ScatterChartStyle6 = 245,
        /// <summary>
        /// Scatter Chart style 7
        /// </summary>
        ScatterChartStyle7 = 246,
        /// <summary>
        /// Scatter Chart style 8
        /// </summary>
        ScatterChartStyle8 = 247,
        /// <summary>
        /// Scatter Chart style 9
        /// </summary>
        ScatterChartStyle9 = 248,
        /// <summary>
        /// Scatter Chart style 10
        /// </summary>
        ScatterChartStyle10 = 249,
        /// <summary>
        /// Scatter Chart style 11
        /// </summary>
        ScatterChartStyle11 = 250,
        /// <summary>
        /// Scatter Chart style 12
        /// </summary>
        ScatterChartStyle12 = 343,
        /// <summary>
        /// Stacked Area 3d Chart Style 1
        /// </summary>
        StackedArea3dChartStyle1 = 276,
        /// <summary>
        /// Stacked Area 3d Chart Style 2
        /// </summary>
        StackedArea3dChartStyle2 = 312,
        /// <summary>
        /// Stacked Area 3d Chart Style 3
        /// </summary>
        StackedArea3dChartStyle3 = 313,
        /// <summary>
        /// Stacked Area 3d Chart Style 4
        /// </summary>
        StackedArea3dChartStyle4 = 314,
        /// <summary>
        /// Stacked Area 3d Chart Style 5
        /// </summary>
        StackedArea3dChartStyle5 = 280,
        /// <summary>
        /// Stacked Area 3d Chart Style 6
        /// </summary>
        StackedArea3dChartStyle6 = 281,
        /// <summary>
        /// Stacked Area 3d Chart Style 7
        /// </summary>
        StackedArea3dChartStyle7 = 282,
        /// <summary>
        /// Stacked Area 3d Chart Style 8
        /// </summary>
        StackedArea3dChartStyle8 = 315,
        /// <summary>
        /// Stacked Area 3d Chart Style 8
        /// </summary>
        StackedArea3dChartStyle9 = 316,
        /// <summary>
        /// Stacked Area 3d Chart Style 10
        /// </summary>
        StackedArea3dChartStyle10 = 350,
        /// <summary>
        /// Stacked Area Chart Style 1
        /// </summary>
        StackedAreaChartStyle1 = 276,
        /// <summary>
        /// Stacked Area Chart Style 2
        /// </summary>
        StackedAreaChartStyle2 = 277,
        /// <summary>
        /// Stacked Area Chart Style 3
        /// </summary>
        StackedAreaChartStyle3 = 278,
        /// <summary>
        /// Stacked Area Chart Style 4
        /// </summary>
        StackedAreaChartStyle4 = 279,
        /// <summary>
        /// Stacked Area Chart Style 5
        /// </summary>
        StackedAreaChartStyle5 = 280,
        /// <summary>
        /// Stacked Area Chart Style 6
        /// </summary>
        StackedAreaChartStyle6 = 281,
        /// <summary>
        /// Stacked Area Chart Style 7
        /// </summary>
        StackedAreaChartStyle7 = 282,
        /// <summary>
        /// Stacked Area Chart Style 8
        /// </summary>
        StackedAreaChartStyle8 = 283,
        /// <summary>
        /// Stacked Area Chart Style 9
        /// </summary>
        StackedAreaChartStyle9 = 284,
        /// <summary>
        /// Stacked Area Chart Style 10
        /// </summary>
        StackedAreaChartStyle10 = 285,
        /// <summary>
        /// Stacked Area Chart Style 11
        /// </summary>
        StackedAreaChartStyle11 = 346,
        /// <summary>
        /// Stacked Column Stacked 3d Chart Style 1
        /// </summary>
        StackedBar3dChartStyle1 = 286,
        /// <summary>
        /// Stacked Column 3d Chart Style 2
        /// </summary>
        StackedBar3dChartStyle2 = 299,
        /// <summary>
        /// Stacked Column 3d Chart Style 3
        /// </summary>
        StackedBar3dChartStyle3 = 310,
        /// <summary>
        /// Stacked Column 3d Chart Style 4
        /// </summary>
        StackedBar3dChartStyle4 = 289,
        /// <summary>
        /// Stacked Column 3d Chart Style 5
        /// </summary>
        StackedBar3dChartStyle5 = 290,
        /// <summary>
        /// Stacked Column 3d Chart Style 6
        /// </summary>
        StackedBar3dChartStyle6 = 294,
        /// <summary>
        /// Stacked Column 3d Chart Style 7
        /// </summary>
        StackedBar3dChartStyle7 = 296,
        /// <summary>
        /// Stacked Column 3d Chart Style 8
        /// </summary>
        StackedBar3dChartStyle8 = 347,
        /// <summary>
        /// Stacked Bar Chart Style 1
        /// </summary>
        StackedBarChartStyle1 = 297,
        /// <summary>
        /// Stacked Bar Chart Style 2
        /// </summary>
        StackedBarChartStyle2 = 298,
        /// <summary>
        /// Stacked Bar Chart Style 3
        /// </summary>
        StackedBarChartStyle3 = 299,
        /// <summary>
        /// Stacked Bar Chart Style 4
        /// </summary>
        StackedBarChartStyle4 = 300,
        /// <summary>
        /// Stacked Bar Chart Style 5
        /// </summary>
        StackedBarChartStyle5 = 301,
        /// <summary>
        /// Stacked Bar Chart Style 6
        /// </summary>
        StackedBarChartStyle6 = 302,
        /// <summary>
        /// Stacked Bar Chart Style 7
        /// </summary>
        StackedBarChartStyle7 = 303,
        /// <summary>
        /// Stacked Bar Chart Style 8
        /// </summary>
        StackedBarChartStyle8 = 304,
        /// <summary>
        /// Stacked Bar Chart Style 9
        /// </summary>
        StackedBarChartStyle9 = 305,
        /// <summary>
        /// Stacked Bar Chart Style 10
        /// </summary>
        StackedBarChartStyle10 = 306,
        /// <summary>
        /// Stacked Bar Chart Style 11
        /// </summary>
        StackedBarChartStyle11 = 348,
        /// <summary>
        /// Stacked Column 3d Chart Style 1
        /// </summary>
        StackedColumn3dChartStyle1 = 286,
        /// <summary>
        /// Stacked Column 3d Chart Style 2
        /// </summary>
        StackedColumn3dChartStyle2 = 299,
        /// <summary>
        /// Stacked Column 3d Chart Style 3
        /// </summary>
        StackedColumn3dChartStyle3 = 310,
        /// <summary>
        /// Stacked Column 3d Chart Style 4
        /// </summary>
        StackedColumn3dChartStyle4 = 289,
        /// <summary>
        /// Stacked Column 3d Chart Style 5
        /// </summary>
        StackedColumn3dChartStyle5 = 290,
        /// <summary>
        /// Stacked Column 3d Chart Style 6
        /// </summary>
        StackedColumn3dChartStyle6 = 294,
        /// <summary>
        /// Stacked Column 3d Chart Style 7
        /// </summary>
        StackedColumn3dChartStyle7 = 296,
        /// <summary>
        /// Stacked Column 3d Chart Style 8
        /// </summary>
        StackedColumn3dChartStyle8 = 347,
        /// <summary>
        /// Stacked Bar Chart style 1
        /// </summary>
        StackedColumnChartStyle1 = 297,
        /// <summary>
        /// Stacked Bar Chart style 2
        /// </summary>
        StackedColumnChartStyle2 = 298,
        /// <summary>
        /// Stacked Bar Chart Style 3
        /// </summary>
        StackedColumnChartStyle3 = 299,
        /// <summary>
        /// Stacked Bar Chart Style 4
        /// </summary>
        StackedColumnChartStyle4 = 300,
        /// <summary>
        /// Stacked Bar Chart Style 5
        /// </summary>
        StackedColumnChartStyle5 = 301,
        /// <summary>
        /// Stacked Bar Chart Style 6
        /// </summary>
        StackedColumnChartStyle6 = 302,
        /// <summary>
        /// Stacked Bar Chart Style 7
        /// </summary>
        StackedColumnChartStyle7 = 303,
        /// <summary>
        /// Stacked Bar Chart Style 8
        /// </summary>
        StackedColumnChartStyle8 = 304,
        /// <summary>
        /// Stacked Bar Chart Style 9
        /// </summary>
        StackedColumnChartStyle9 = 305,
        /// <summary>
        /// Stacked Bar Chart Style 10
        /// </summary>
        StackedColumnChartStyle10 = 306,
        /// <summary>
        /// Stacked Bar Chart Style 11
        /// </summary>
        StackedColumnChartStyle11 = 348,
    }
    /// <summary>
    /// Maps to Excel's built-in chart styles, for charts with more that one data serie. 
    /// Note that Excel changes chart type depending on many parameters, like number of series, axis type and more, so it will not always match the number in Excel.
    /// To be certain of getting the correct style use the chart style number of the style you want to apply
    /// For charts with only one data serie use <see cref="ePresetChartStyle"/>
    /// Styles are fetched from the StyleLibrary by the id provided in this enum. 
    /// <seealso cref="ExcelChartStyleManager.StyleLibrary" />
    /// </summary>
    public enum ePresetChartStyleMultiSeries
    {
        /// <summary>
        /// 3D Area Chart style 1
        /// </summary>
        Area3dChartStyle1 = 276,
        /// <summary>
        /// 3D Area Chart style 2
        /// </summary>
        Area3dChartStyle2 = 312,
        /// <summary>
        /// 3D Area Chart style 3
        /// </summary>
        Area3dChartStyle3 = 313,
        /// <summary>
        /// 3D Area Chart style 4
        /// </summary>
        Area3dChartStyle4 = 314,
        /// <summary>
        /// 3D Area Chart style 5
        /// </summary>
        Area3dChartStyle5 = 280,
        /// <summary>
        /// 3D Area Chart style 6
        /// </summary>
        Area3dChartStyle6 = 281,
        /// <summary>
        /// 3D Area Chart style 7
        /// </summary>
        Area3dChartStyle7 = 282,
        /// <summary>
        /// 3D Area Chart style 8
        /// </summary>
        Area3dChartStyle8 = 315,
        /// <summary>
        /// 3D Area Chart style 9
        /// </summary>
        Area3dChartStyle9 = 350,
        /// <summary>
        /// Area Chart style 1
        /// </summary>
        AreaChartStyle1 = 276,
        /// <summary>
        /// Area Chart style 2
        /// </summary>
        AreaChartStyle2 = 277,
        /// <summary>
        /// Area Chart style 3
        /// </summary>
        AreaChartStyle3 = 278,
        /// <summary>
        /// Area Chart style 4
        /// </summary>
        AreaChartStyle4 = 279,
        /// <summary>
        /// Area Chart style 5
        /// </summary>
        AreaChartStyle5 = 280,
        /// <summary>
        /// Area Chart style 6
        /// </summary>
        AreaChartStyle6 = 281,
        /// <summary>
        /// Area Chart style 7
        /// </summary>
        AreaChartStyle7 = 282,
        /// <summary>
        /// Area Chart style 8
        /// </summary>
        AreaChartStyle8 = 283,
        /// <summary>
        /// Area Chart style 9
        /// </summary>
        AreaChartStyle9 = 284,
        /// <summary>
        /// Area Chart style 10
        /// </summary>
        AreaChartStyle10 = 346,
        /// <summary>
        /// Bar 3d Chart Style 1
        /// </summary>
        Bar3dChartStyle1 = 286,
        /// <summary>
        /// Bar 3d Chart Style 2
        /// </summary>
        Bar3dChartStyle2 = 287,
        /// <summary>
        /// Bar 3d Chart Style 3
        /// </summary>
        Bar3dChartStyle3 = 288,
        /// <summary>
        /// Bar 3d Chart Style 4
        /// </summary>
        Bar3dChartStyle4 = 289,
        /// <summary>
        /// Bar 3d Chart Style 5
        /// </summary>
        Bar3dChartStyle5 = 290,
        /// <summary>
        /// Bar 3d Chart Style 6
        /// </summary>
        Bar3dChartStyle6 = 291,
        /// <summary>
        /// Bar 3d Chart Style 7
        /// </summary>
        Bar3dChartStyle7 = 292,
        /// <summary>
        /// Bar 3d Chart Style 8
        /// </summary>
        Bar3dChartStyle8 = 349,
        /// <summary>
        /// Bar 3d Chart Style 9
        /// </summary>
        Bar3dChartStyle9 = 294,
        /// <summary>
        /// Bar 3d Chart Style 10
        /// </summary>
        Bar3dChartStyle10 = 296,
        /// <summary>
        /// Bar 3d Chart Style 11
        /// </summary>
        Bar3dChartStyle11 = 347,
        /// <summary>
        /// Bar Chart style 1
        /// </summary>
        BarChartStyle1 = 216,
        /// <summary>
        /// Bar Chart style 2
        /// </summary>
        BarChartStyle2 = 217,
        /// <summary>
        /// Bar Chart style 3
        /// </summary>
        BarChartStyle3 = 218,
        /// <summary>
        /// Bar Chart style 4
        /// </summary>
        BarChartStyle4 = 219,
        /// <summary>
        /// Bar Chart style 5
        /// </summary>
        BarChartStyle5 = 220,
        /// <summary>
        /// Bar Chart style 6
        /// </summary>
        BarChartStyle6 = 221,
        /// <summary>
        /// Bar Chart style 7
        /// </summary>
        BarChartStyle7 = 222,
        /// <summary>
        /// Bar Chart style 8
        /// </summary>
        BarChartStyle8 = 223,
        /// <summary>
        /// Bar Chart style 9
        /// </summary>
        BarChartStyle9 = 224,
        /// <summary>
        /// Bar Chart style 10
        /// </summary>
        BarChartStyle10 = 225,
        /// <summary>
        /// Bar Chart style 11
        /// </summary>
        BarChartStyle11 = 339,
        /// <summary>
        /// Bar Chart style 12
        /// </summary>
        BarChartStyle12 = 341,
        /// <summary>
        /// Bubble 3d Chart Style 1
        /// </summary>
        Bubble3dChartStyle1 = 269,
        /// <summary>
        /// Bubble 3d Chart Style 2
        /// </summary>
        Bubble3dChartStyle2 = 270,
        /// <summary>
        /// Bubble 3d Chart Style 3
        /// </summary>
        Bubble3dChartStyle3 = 272,
        /// <summary>
        /// Bubble 3d Chart Style 4
        /// </summary>
        Bubble3dChartStyle4 = 246,
        /// <summary>
        /// Bubble 3d Chart Style 5
        /// </summary>
        Bubble3dChartStyle5 = 242,
        /// <summary>
        /// Bubble 3d Chart Style 6
        /// </summary>
        Bubble3dChartStyle6 = 273,
        /// <summary>
        /// Bubble 3d Chart Style 7
        /// </summary>
        Bubble3dChartStyle7 = 248,
        /// <summary>
        /// Bubble 3d Chart Style 8
        /// </summary>
        Bubble3dChartStyle8 = 343,
        /// <summary>
        /// Bubble Chart Style 1
        /// </summary>
        BubbleChartStyle1 = 269,
        /// <summary>
        /// Bubble Chart Style 2
        /// </summary>
        BubbleChartStyle2 = 270,
        /// <summary>
        /// Bubble Chart Style 3
        /// </summary>
        BubbleChartStyle3 = 271,
        /// <summary>
        /// Bubble Chart Style 4
        /// </summary>
        BubbleChartStyle4 = 272,
        /// <summary>
        /// Bubble Chart Style 5
        /// </summary>
        BubbleChartStyle5 = 246,
        /// <summary>
        /// Bubble Chart Style 6
        /// </summary>
        BubbleChartStyle6 = 242,
        /// <summary>
        /// Bubble Chart Style 7
        /// </summary>
        BubbleChartStyle7 = 273,
        /// <summary>
        /// Bubble Chart Style 8
        /// </summary>
        BubbleChartStyle8 = 248,
        /// <summary>
        /// Bubble Chart Style 9
        /// </summary>
        BubbleChartStyle9 = 274,
        /// <summary>
        /// Bubble Chart Style 10
        /// </summary>
        BubbleChartStyle10 = 343,
        /// <summary>
        /// Column Chart style 1
        /// </summary>
        ColumnChartStyle1 = 201,
        /// <summary>
        /// Column Chart style 2
        /// </summary>
        ColumnChartStyle2 = 202,
        /// <summary>
        /// Column Chart style 3
        /// </summary>
        ColumnChartStyle3 = 203,
        /// <summary>
        /// Column Chart style 4
        /// </summary>
        ColumnChartStyle4 = 205,
        /// <summary>
        /// Column Chart style 5
        /// </summary>
        ColumnChartStyle5 = 206,
        /// <summary>
        /// Column Chart style 6
        /// </summary>
        ColumnChartStyle6 = 207,
        /// <summary>
        /// Column Chart style 7
        /// </summary>
        ColumnChartStyle7 = 208,
        /// <summary>
        /// Column Chart style 8
        /// </summary>
        ColumnChartStyle8 = 209,
        /// <summary>
        /// Column Chart style 9
        /// </summary>
        ColumnChartStyle9 = 210,
        /// <summary>
        /// Column Chart style 10
        /// </summary>
        ColumnChartStyle10 = 211,
        /// <summary>
        /// Column Chart style 11
        /// </summary>
        ColumnChartStyle11 = 212,
        /// <summary>
        /// Column Chart style 12
        /// </summary>
        ColumnChartStyle12 = 213,
        /// <summary>
        /// Column Chart style 13
        /// </summary>
        ColumnChartStyle13 = 215,
        /// <summary>
        /// Column Chart style 14
        /// </summary>
        ColumnChartStyle14 = 340,
        /// <summary>
        /// Column 3d Chart Style 1
        /// </summary>
        Column3dChartStyle1 = 286,
        /// <summary>
        /// Column 3d Chart Style 2
        /// </summary>
        Column3dChartStyle2 = 287,
        /// <summary>
        /// Column 3d Chart Style 3
        /// </summary>
        Column3dChartStyle3 = 288,
        /// <summary>
        /// Column 3d Chart Style 4
        /// </summary>
        Column3dChartStyle4 = 289,
        /// <summary>
        /// Column 3d Chart Style 5
        /// </summary>
        Column3dChartStyle5 = 290,
        /// <summary>
        /// Column 3d Chart Style 6
        /// </summary>
        Column3dChartStyle6 = 291,
        /// <summary>
        /// Column 3d Chart Style 7
        /// </summary>
        Column3dChartStyle7 = 292,
        /// <summary>
        /// Column 3d Chart Style 8
        /// </summary>
        Column3dChartStyle8 = 293,
        /// <summary>
        /// Column 3d Chart Style 9
        /// </summary>
        Column3dChartStyle9 = 294,
        /// <summary>
        /// Column 3d Chart Style 10
        /// </summary>
        Column3dChartStyle10 = 296,
        /// <summary>
        /// Column 3d Chart Style 11
        /// </summary>
        Column3dChartStyle11 = 347,
        /// <summary>
        /// Custom Combined Chart Style 1
        /// </summary>
        ComboChartStyle1 = 322,
        /// <summary>
        /// Custom Combined Chart Style 2
        /// </summary>
        ComboChartStyle2 = 323,
        /// <summary>
        /// Custom Combined Chart Style 3
        /// </summary>
        ComboChartStyle3 = 325,
        /// <summary>
        /// Custom Combined Chart Style 4
        /// </summary>
        ComboChartStyle4 = 326,
        /// <summary>
        /// Custom Combined Chart Style 5
        /// </summary>
        ComboChartStyle5 = 221,
        /// <summary>
        /// Custom Combined Chart Style 6
        /// </summary>
        ComboChartStyle6 = 328,
        /// <summary>
        /// Custom Combined Chart Style 7
        /// </summary>
        ComboChartStyle7 = 225,
        /// <summary>
        /// Custom Combined Chart Style 8
        /// </summary>
        ComboChartStyle8 = 352,
        /// <summary>
        /// Doughnut Chart Style 1
        /// </summary>
        DoughnutChartStyle1 = 251,
        /// <summary>
        /// Doughnut Chart Style 2
        /// </summary>
        DoughnutChartStyle2 = 252,
        /// <summary>
        /// Doughnut Chart Style 3
        /// </summary>
        DoughnutChartStyle3 = 253,
        /// <summary>
        /// Doughnut Chart Style 4
        /// </summary>
        DoughnutChartStyle4 = 254,
        /// <summary>
        /// Doughnut Chart Style 5
        /// </summary>
        DoughnutChartStyle5 = 255,
        /// <summary>
        /// Doughnut Chart Style 6
        /// </summary>
        DoughnutChartStyle6 = 256,
        /// <summary>
        /// Doughnut Chart Style 7
        /// </summary>
        DoughnutChartStyle7 = 257,
        /// <summary>
        /// Doughnut Chart Style 8
        /// </summary>
        DoughnutChartStyle8 = 258,
        /// <summary>
        /// Doughnut Chart Style 9
        /// </summary>
        DoughnutChartStyle9 = 261,
        /// <summary>
        /// Line 3d Chart style 1
        /// </summary>
        Line3dChartStyle1 = 307,
        /// <summary>
        /// Line 3d Chart style 2
        /// </summary>
        Line3dChartStyle2 = 311,
        /// <summary>
        /// Line 3d Chart style 3
        /// </summary>
        Line3dChartStyle3 = 308,
        /// <summary>
        /// Line 3d Chart style 4
        /// </summary>
        Line3dChartStyle4 = 309,
        /// <summary>
        /// Line Chart style 1
        /// </summary>
        LineChartStyle1 = 227,
        /// <summary>
        /// Line Chart style 2
        /// </summary>
        LineChartStyle2 = 228,
        /// <summary>
        /// Line Chart style 3
        /// </summary>
        LineChartStyle3 = 230,
        /// <summary>
        /// Line Chart style 4
        /// </summary>
        LineChartStyle4 = 231,
        /// <summary>
        /// Line Chart style 5
        /// </summary>
        LineChartStyle5 = 232,
        /// <summary>
        /// Line Chart style 6
        /// </summary>
        LineChartStyle6 = 233,
        /// <summary>
        /// Line Chart style 7
        /// </summary>
        LineChartStyle7 = 234,
        /// <summary>
        /// Line Chart style 8
        /// </summary>
        LineChartStyle8 = 235,
        /// <summary>
        /// Line Chart style 9
        /// </summary>
        LineChartStyle9 = 236,
        /// <summary>
        /// Line Chart style 10
        /// </summary>
        LineChartStyle10 = 237,
        /// <summary>
        /// Line Chart style 11
        /// </summary>
        LineChartStyle11 = 239,
        /// <summary>
        /// Line Chart style 12
        /// </summary>
        LineChartStyle12 = 332,
        /// <summary>
        /// Line Chart style 13
        /// </summary>
        LineChartStyle13 = 342,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 1
        /// </summary>
        OfPieChartStyle1 = 333,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 2
        /// </summary>
        OfPieChartStyle2 = 252,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 3
        /// </summary>
        OfPieChartStyle3 = 334,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 4
        /// </summary>
        OfPieChartStyle4 = 335,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 5
        /// </summary>
        OfPieChartStyle5 = 336,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 6
        /// </summary>
        OfPieChartStyle6 = 337,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 7
        /// </summary>
        OfPieChartStyle7 = 338,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 8
        /// </summary>
        OfPieChartStyle8 = 258,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 9
        /// </summary>
        OfPieChartStyle9 = 259,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 10
        /// </summary>
        OfPieChartStyle10 = 260,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 11
        /// </summary>
        OfPieChartStyle11 = 261,
        /// <summary>
        /// Pie- or Bar-of pie Chart style 12
        /// </summary>
        OfPieChartStyle12 = 344,
        /// <summary>
        /// Pie Chart Style 1
        /// </summary>
        PieChartStyle1 = 251,
        /// <summary>
        /// Pie Chart Style 2
        /// </summary>
        PieChartStyle2 = 252,
        /// <summary>
        /// Pie Chart Style 3
        /// </summary>
        PieChartStyle3 = 253,
        /// <summary>
        /// Pie Chart Style 4
        /// </summary>
        PieChartStyle4 = 254,
        /// <summary>
        /// Pie Chart Style 5
        /// </summary>
        PieChartStyle5 = 255,
        /// <summary>
        /// Pie Chart Style 6
        /// </summary>
        PieChartStyle6 = 256,
        /// <summary>
        /// Pie Chart Style 7
        /// </summary>
        PieChartStyle7 = 257,
        /// <summary>
        /// Pie Chart Style 8
        /// </summary>
        PieChartStyle8 = 258,
        /// <summary>
        /// Pie Chart Style 9
        /// </summary>
        PieChartStyle9 = 259,
        /// <summary>
        /// Pie Chart Style 10
        /// </summary>
        PieChartStyle10 = 260,
        /// <summary>
        /// Pie Chart style 11
        /// </summary>
        PieChartStyle11 = 261,
        /// <summary>
        /// Pie Chart style 12
        /// </summary>
        PieChartStyle12 = 344,
        /// <summary>
        /// Pie 3d Chart Style 1
        /// </summary>
        Pie3dChartStyle1 = 262,
        /// <summary>
        /// Pie 3d Chart Style 2
        /// </summary>
        Pie3dChartStyle2 = 263,
        /// <summary>
        /// Pie 3d Chart Style 3
        /// </summary>
        Pie3dChartStyle3 = 264,
        /// <summary>
        /// Pie 3d Chart Style 4
        /// </summary>
        Pie3dChartStyle4 = 265,
        /// <summary>
        /// Pie 3d Chart Style 5
        /// </summary>
        Pie3dChartStyle5 = 266,
        /// <summary>
        /// Pie 3d Chart Style 6
        /// </summary>
        Pie3dChartStyle6 = 267,
        /// <summary>
        /// Pie 3d Chart Style 7
        /// </summary>
        Pie3dChartStyle7 = 268,
        /// <summary>
        /// Pie 3d Chart Style 8
        /// </summary>
        Pie3dChartStyle8 = 259,
        /// <summary>
        /// Pie 3d Chart Style 9
        /// </summary>
        Pie3dChartStyle9 = 261,
        /// <summary>
        /// Pie 3d Chart Style 10
        /// </summary>
        Pie3dChartStyle10 = 345,
        /// <summary>
        /// Radar Chart style 1
        /// </summary>
        RadarChartStyle1 = 317,
        /// <summary>
        /// Radar Chart style 2
        /// </summary>
        RadarChartStyle2 = 318,
        /// <summary>
        /// Radar Chart style 3
        /// </summary>
        RadarChartStyle3 = 206,
        /// <summary>
        /// Radar Chart style 4
        /// </summary>
        RadarChartStyle4 = 319,
        /// <summary>
        /// Radar Chart style 5
        /// </summary>
        RadarChartStyle5 = 320,
        /// <summary>
        /// Radar Chart style 6
        /// </summary>
        RadarChartStyle6 = 207,
        /// <summary>
        /// Radar Chart style 7
        /// </summary>
        RadarChartStyle7 = 321,
        /// <summary>
        /// Radar Chart style 8
        /// </summary>
        RadarChartStyle8 = 351,
        /// <summary>
        /// Scatter Chart style 1
        /// </summary>
        ScatterChartStyle1 = 240,
        /// <summary>
        /// Scatter Chart style 2
        /// </summary>
        ScatterChartStyle2 = 241,
        /// <summary>
        /// Scatter Chart style 3
        /// </summary>
        ScatterChartStyle3 = 242,
        /// <summary>
        /// Scatter Chart style 4
        /// </summary>
        ScatterChartStyle4 = 243,
        /// <summary>
        /// Scatter Chart style 5
        /// </summary>
        ScatterChartStyle5 = 244,
        /// <summary>
        /// Scatter Chart style 6
        /// </summary>
        ScatterChartStyle6 = 245,
        /// <summary>
        /// Scatter Chart style 7
        /// </summary>
        ScatterChartStyle7 = 246,
        /// <summary>
        /// Scatter Chart style 8
        /// </summary>
        ScatterChartStyle8 = 248,
        /// <summary>
        /// Scatter Chart style 9
        /// </summary>
        ScatterChartStyle9 = 249,
        /// <summary>
        /// Scatter Chart style 10
        /// </summary>
        ScatterChartStyle10 = 250,
        /// <summary>
        /// Scatter Chart style 11
        /// </summary>
        ScatterChartStyle11 = 343,
        /// <summary>
        /// Stacked Area 3d Chart Style 1
        /// </summary>
        StackedArea3dChartStyle1 = 276,
        /// <summary>
        /// Stacked Area 3d Chart Style 2
        /// </summary>
        StackedArea3dChartStyle2 = 312,
        /// <summary>
        /// Stacked Area 3d Chart Style 3
        /// </summary>
        StackedArea3dChartStyle3 = 313,
        /// <summary>
        /// Stacked Area 3d Chart Style 4
        /// </summary>
        StackedArea3dChartStyle4 = 314,
        /// <summary>
        /// Stacked Area 3d Chart Style 5
        /// </summary>
        StackedArea3dChartStyle5 = 280,
        /// <summary>
        /// Stacked Area 3d Chart Style 6
        /// </summary>
        StackedArea3dChartStyle6 = 281,
        /// <summary>
        /// Stacked Area 3d Chart Style 7
        /// </summary>
        StackedArea3dChartStyle7 = 282,
        /// <summary>
        /// Stacked Area 3d Chart Style 8
        /// </summary>
        StackedArea3dChartStyle8 = 315,
        /// <summary>
        /// Stacked Area 3d Chart Style 9
        /// </summary>
        StackedArea3dChartStyle9 = 350,
        /// <summary>
        /// Stacked Area Chart Style 1
        /// </summary>
        StackedAreaChartStyle1 = 276,
        /// <summary>
        /// Stacked Area Chart Style 2
        /// </summary>
        StackedAreaChartStyle2 = 277,
        /// <summary>
        /// Stacked Area Chart Style 3
        /// </summary>
        StackedAreaChartStyle3 = 278,
        /// <summary>
        /// Stacked Area Chart Style 4
        /// </summary>
        StackedAreaChartStyle4 = 279,
        /// <summary>
        /// Stacked Area Chart Style 5
        /// </summary>
        StackedAreaChartStyle5 = 280,
        /// <summary>
        /// Stacked Area Chart Style 6
        /// </summary>
        StackedAreaChartStyle6 = 281,
        /// <summary>
        /// Stacked Area Chart Style 7
        /// </summary>
        StackedAreaChartStyle7 = 282,
        /// <summary>
        /// Stacked Area Chart Style 8
        /// </summary>
        StackedAreaChartStyle8 = 283,
        /// <summary>
        /// Stacked Area Chart Style 9
        /// </summary>
        StackedAreaChartStyle9 = 284,
        /// <summary>
        /// Stacked Area Chart Style 10
        /// </summary>
        StackedAreaChartStyle10 = 346,
        /// <summary>
        /// Stacked Column Stacked 3d Chart Style 1
        /// </summary>
        StackedBar3dChartStyle1 = 286,
        /// <summary>
        /// Stacked Column 3d Chart Style 2
        /// </summary>
        StackedBar3dChartStyle2 = 299,
        /// <summary>
        /// Stacked Column 3d Chart Style 3
        /// </summary>
        StackedBar3dChartStyle3 = 310,
        /// <summary>
        /// Stacked Column 3d Chart Style 4
        /// </summary>
        StackedBar3dChartStyle4 = 289,
        /// <summary>
        /// Stacked Column 3d Chart Style 5
        /// </summary>
        StackedBar3dChartStyle5 = 290,
        /// <summary>
        /// Stacked Column 3d Chart Style 6
        /// </summary>
        StackedBar3dChartStyle6 = 294,
        /// <summary>
        /// Stacked Column 3d Chart Style 7
        /// </summary>
        StackedBar3dChartStyle7 = 296,
        /// <summary>
        /// Stacked Column 3d Chart Style 8
        /// </summary>
        StackedBar3dChartStyle8 = 347,
        /// <summary>
        /// Stacked Bar Chart Style 1
        /// </summary>
        StackedBarChartStyle1 = 297,
        /// <summary>
        /// Stacked Bar Chart Style 2
        /// </summary>
        StackedBarChartStyle2 = 298,
        /// <summary>
        /// Stacked Bar Chart Style 3
        /// </summary>
        StackedBarChartStyle3 = 299,
        /// <summary>
        /// Stacked Bar Chart Style 4
        /// </summary>
        StackedBarChartStyle4 = 300,
        /// <summary>
        /// Stacked Bar Chart Style 5
        /// </summary>
        StackedBarChartStyle5 = 301,
        /// <summary>
        /// Stacked Bar Chart Style 6
        /// </summary>
        StackedBarChartStyle6 = 302,
        /// <summary>
        /// Stacked Bar Chart Style 7
        /// </summary>
        StackedBarChartStyle7 = 303,
        /// <summary>
        /// Stacked Bar Chart Style 8
        /// </summary>
        StackedBarChartStyle8 = 304,
        /// <summary>
        /// Stacked Bar Chart Style 9
        /// </summary>
        StackedBarChartStyle9 = 305,
        /// <summary>
        /// Stacked Bar Chart Style 10
        /// </summary>
        StackedBarChartStyle10 = 306,
        /// <summary>
        /// Stacked Bar Chart Style 11
        /// </summary>
        StackedBarChartStyle11 = 348,
        /// <summary>
        /// Stacked Column 3d Chart Style 1
        /// </summary>
        StackedColumn3dChartStyle1 = 286,
        /// <summary>
        /// Stacked Column 3d Chart Style 2
        /// </summary>
        StackedColumn3dChartStyle2 = 299,
        /// <summary>
        /// Stacked Column 3d Chart Style 3
        /// </summary>
        StackedColumn3dChartStyle3 = 310,
        /// <summary>
        /// Stacked Column 3d Chart Style 4
        /// </summary>
        StackedColumn3dChartStyle4 = 289,
        /// <summary>
        /// Stacked Column 3d Chart Style 5
        /// </summary>
        StackedColumn3dChartStyle5 = 290,
        /// <summary>
        /// Stacked Column 3d Chart Style 6
        /// </summary>
        StackedColumn3dChartStyle6 = 294,
        /// <summary>
        /// Stacked Column 3d Chart Style 7
        /// </summary>
        StackedColumn3dChartStyle7 = 296,
        /// <summary>
        /// Stacked Column 3d Chart Style 8
        /// </summary>
        StackedColumn3dChartStyle8 = 347,
        /// <summary>
        /// Stacked Bar Chart style 1
        /// </summary>
        StackedColumnChartStyle1 = 297,
        /// <summary>
        /// Stacked Bar Chart style 2
        /// </summary>
        StackedColumnChartStyle2 = 298,
        /// <summary>
        /// Stacked Bar Chart Style 3
        /// </summary>
        StackedColumnChartStyle3 = 299,
        /// <summary>
        /// Stacked Bar Chart Style 4
        /// </summary>
        StackedColumnChartStyle4 = 300,
        /// <summary>
        /// Stacked Bar Chart Style 5
        /// </summary>
        StackedColumnChartStyle5 = 301,
        /// <summary>
        /// Stacked Bar Chart Style 6
        /// </summary>
        StackedColumnChartStyle6 = 302,
        /// <summary>
        /// Stacked Bar Chart Style 7
        /// </summary>
        StackedColumnChartStyle7 = 303,
        /// <summary>
        /// Stacked Bar Chart Style 8
        /// </summary>
        StackedColumnChartStyle8 = 304,
        /// <summary>
        /// Stacked Bar Chart Style 9
        /// </summary>
        StackedColumnChartStyle9 = 305,
        /// <summary>
        /// Stacked Bar Chart Style 10
        /// </summary>
        StackedColumnChartStyle10 = 306,
        /// <summary>
        /// Stacked Bar Chart Style 11
        /// </summary>
        StackedColumnChartStyle11 = 348,
    }
    /// <summary>
    /// Chart color schemes mapping to the default colors in Excel
    /// </summary>
    public enum ePresetChartColors
    {
        /// <summary>
        /// Colorful Palette 1
        /// </summary>
        ColorfulPalette1 = 10,
        /// <summary>
        /// Colorful Palette 2
        /// </summary>
        ColorfulPalette2 = 11,
        /// <summary>
        /// Colorful Palette 3
        /// </summary>
        ColorfulPalette3 = 12,
        /// <summary>
        /// Colorful Palette 4
        /// </summary>
        ColorfulPalette4 = 13,
        /// <summary>
        /// Monochromatic Palette 1
        /// </summary>
        MonochromaticPalette1 = 14,
        /// <summary>
        /// Monochromatic Palette 2
        /// </summary>
        MonochromaticPalette2 = 15,
        /// <summary>
        /// Monochromatic Palette 3
        /// </summary>
        MonochromaticPalette3 = 16,
        /// <summary>
        /// Monochromatic Palette 4
        /// </summary>
        MonochromaticPalette4 = 17,
        /// <summary>
        /// Monochromatic Palette 5
        /// </summary>
        MonochromaticPalette5 = 18,
        /// <summary>
        /// Monochromatic Palette 6
        /// </summary>
        MonochromaticPalette6 = 19,
        /// <summary>
        /// Monochromatic Palette 7
        /// </summary>
        MonochromaticPalette7 = 20,
        /// <summary>
        /// Monochromatic Palette 8
        /// </summary>
        MonochromaticPalette8 = 21,
        /// <summary>
        /// Monochromatic Palette 9
        /// </summary>
        MonochromaticPalette9 = 22,
        /// <summary>
        /// Monochromatic Palette 10
        /// </summary>
        MonochromaticPalette10 = 23,
        /// <summary>
        /// Monochromatic Palette 11
        /// </summary>
        MonochromaticPalette11 = 24,
        /// <summary>
        /// Monochromatic Palette 12
        /// </summary>
        MonochromaticPalette12 = 25,
        /// <summary>
        /// Monochromatic Palette 13
        /// </summary>
        MonochromaticPalette13 = 26
    }
}