/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Chart.ChartEx   
{
    public enum eProjectionType
    {
        /// <summary>
        /// Automatic
        /// </summary>
        Automatic,
        /// <summary>
        /// A Mercator projection.
        /// </summary>
        Mercator,
        /// <summary>
        /// a Miller cylindrical projection.
        /// </summary>
        Miller,
        /// <summary>
        /// A Robinson projection.
        /// </summary>
        Robinson,
        /// <summary>
        /// An Albers equal-area conic projection.
        /// </summary>
        Albers
    }
}