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
    /// Position of the labels
    /// </summary>
    public enum eLabelPosition
    {
        /// <summary>
        /// Best fit
        /// </summary>
        BestFit,
        /// <summary>
        /// Left aligned
        /// </summary>
        Left,
        /// <summary>
        /// Right aligned
        /// </summary>
        Right,
        /// <summary>
        /// Center aligned
        /// </summary>
        Center,
        /// <summary>
        /// Top aligned
        /// </summary>
        Top,
        /// <summary>
        /// Bottom aligned
        /// </summary>
        Bottom,
        /// <summary>
        /// Labels will be displayed inside the data marker
        /// </summary>
        InBase,
        /// <summary>
        /// Labels will be displayed inside the end of the data marker
        /// </summary>
        InEnd,
        /// <summary>
        /// Labels will be displayed outside the end of the data marker
        /// </summary>
        OutEnd
    }
}