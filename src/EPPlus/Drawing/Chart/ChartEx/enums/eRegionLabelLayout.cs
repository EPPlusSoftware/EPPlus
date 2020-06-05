/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    04/15/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// The layout type for region labels of a geospatial series
    /// </summary>
    public enum eRegionLabelLayout
    {
        /// <summary>
        /// No region labels appear in a geospatial series
        /// </summary>
        None,
        /// <summary>
        /// Region labels only appear if they can fit in their respective containing geometries in a geospatial series
        /// </summary>
        BestFitOnly,
        /// <summary>
        /// All region labels appear
        /// </summary>
        All
    }
}
