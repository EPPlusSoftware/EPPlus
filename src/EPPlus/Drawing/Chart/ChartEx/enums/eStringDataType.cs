/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// Side positions for a chart element
    /// </summary>
    public enum eStringDataType
    {
        /// <summary>
        /// The category string dimension data type.
        /// </summary>
        Category,
        /// <summary>
        /// The string dimension associated with a color.
        /// </summary>
        ColorString,
        /// <summary>
        /// The geographical entity identifier string dimension data type. 
        /// This dimension can be used to provide locations to a geospatial series in a Geographic chart. 
        /// Refer to the usage of entityId in Geo Cache and Data.
        /// </summary>
        EntityId
    }
}