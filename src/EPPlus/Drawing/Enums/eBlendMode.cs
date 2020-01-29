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
    /// How to render effects one on top of another
    /// </summary>
    public enum eBlendMode
    {
        /// <summary>
        /// Overlay
        /// </summary>
        Over,
        /// <summary>
        /// Multiply
        /// </summary>
        Mult,
        /// <summary>
        /// Screen
        /// </summary>
        Screen,
        /// <summary>
        /// Darken
        /// </summary>
        Darken,
        /// <summary>
        /// Lighten
        /// </summary>
        Lighten
    }
}