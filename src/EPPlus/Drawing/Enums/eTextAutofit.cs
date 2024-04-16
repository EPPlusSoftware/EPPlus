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
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// How autofit handles text.
    /// </summary>
    public enum eTextAutofit
    {
        /// <summary>
        /// No autofit
        /// </summary>
        NoAutofit,
        /// <summary>
        /// Text within the text body will be normally autofit
        /// </summary>
        NormalAutofit,
        /// <summary>
        /// A shape will be autofit to fully contain the text within it
        /// </summary>
        ShapeAutofit
    }
}