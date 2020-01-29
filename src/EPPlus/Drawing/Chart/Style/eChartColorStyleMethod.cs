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
    /// Method for how colors are picked from the Colors collection
    /// </summary>
    public enum eChartColorStyleMethod
    {
        /// <summary>
        /// The color picked from Colors is the index modulus the total set of colors in the list.
        /// </summary>
        Cycle,
        /// <summary>
        /// The color picked from Colors is the first color with a brightness that varies from darker to lighter.
        /// </summary>
        WithinLinear,
        /// <summary>
        /// The color picked from Colors is the index modulus the total set of colors in the list. The brightness varies from lighter to darker
        /// </summary>
        AcrossLinear,
        /// <summary>
        /// The color picked from Colors is the first color with a brightness that varies from lighter to darker. The brightness varies from darker to lighter. 
        /// </summary>
        WithinLinearReversed,
        /// <summary>
        /// The color picked from Colors is the index modulus the total set of colors in the list. The brightness varies from darkerlighter. 
        /// </summary>
        AcrossLinearReversed
    }
}