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
    /// Bar chart type
    /// </summary>
    public enum eBarChartType
    {
        /// <summary>
        /// A clustered 3D bar chart
        /// </summary>
        BarClustered3D = 60,
        /// <summary>
        /// A stacked 3D bar chart
        /// </summary>
        BarStacked3D = 61,
        /// <summary>
        /// A Stacked 100 percent 3D bar chart
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
        /// A stacked 100 percent 3D column chart
        /// </summary>
        ColumnStacked1003D = 56,
        /// <summary>
        /// A clustered bar chart
        /// </summary>
        BarClustered = 57,
        /// <summary>
        /// A stacked bar chart
        /// </summary>
        BarStacked = 58,
        /// <summary>
        /// A stacked 100 percent bar chart
        /// </summary>
        BarStacked100 = 59,
        /// <summary>
        /// A clustered column chart 
        /// </summary>
        ColumnClustered = 51,
        /// <summary>
        /// A stacked column chart
        /// </summary>
        ColumnStacked = 52,
        /// <summary>
        /// A stacked column 100 percent chart
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
        /// A stacked 100 percent cone bar chart
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
        /// A stacked 100 percent cone column chart 
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
        /// A stacked 100 percent cylinder bar chart
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
        /// A stacked 100 percent cylinder column chart
        /// </summary>
        CylinderColStacked100 = 94,
        /// <summary>
        /// A clustered pyramid bar chart
        /// </summary>
        PyramidBarClustered = 109,
        /// <summary>
        /// A stacked pyramid bar chart
        /// </summary>
        PyramidBarStacked = 110,
        /// <summary>
        /// A stacked 100 percent pyramid bar chart
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
        /// A stacked 100 percent pyramid column chart
        /// </summary>
        PyramidColStacked100 = 108
    }
}