/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/23/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// The line style of a vml drawing
    /// </summary>
    public enum eVmlLineStyle
    {
        /// <summary>
        /// No line style
        /// </summary>
        NoLine,
        /// <summary>
        /// A single line
        /// </summary>
        Single,
        /// <summary>
        /// Thin thin line style
        /// </summary>
        ThinThin,
        /// <summary>
        /// Thin thick line style
        /// </summary>
        ThinThick,
        /// <summary>
        /// Thick thin line style
        /// </summary>
        ThickThin,
        /// <summary>
        /// Thick between thin line style
        /// </summary>
        ThickBetweenThin
    }
}