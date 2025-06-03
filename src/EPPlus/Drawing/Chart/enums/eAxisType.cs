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
    /// Axis type
    /// </summary>
    public enum eAxisType
    {
        /// <summary>
        /// Value axis
        /// </summary>
        Val,
        /// <summary>
        /// Category axis
        /// </summary>
        Cat,
        /// <summary>
        /// Date axis
        /// </summary>
        Date,
        /// <summary>
        /// Series axis (Type of Category axis usually in 3D charts)
        /// </summary>
        Serie
    }
}