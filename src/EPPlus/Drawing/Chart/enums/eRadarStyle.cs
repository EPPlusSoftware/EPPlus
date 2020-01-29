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
    /// Radar chart type
    /// </summary>
    public enum eRadarStyle
    {
        /// <summary>
        /// The radar chart will be filled and have lines, but will not have markers.
        /// </summary>
        Filled,
        /// <summary>
        /// The radar chart will have lines and markers, but will not be filled.
        /// </summary>
        Marker,
        /// <summary>
        /// The radar chart will have lines, but no markers and no filling.
        /// </summary>
        Standard 
    }
}