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
    /// Line end style.
    /// </summary>
    public enum eEndStyle   //ST_LineEndType
    {
        /// <summary>
        /// No end
        /// </summary>
        None,
        /// <summary>
        /// Triangle arrow head
        /// </summary>
        Triangle,
        /// <summary>
        /// Stealth arrow head
        /// </summary>
        Stealth,
        /// <summary>
        /// Diamond
        /// </summary>
        Diamond,
        /// <summary>
        /// Oval
        /// </summary>
        Oval,
        /// <summary>
        /// Line arrow head
        /// </summary>
        Arrow
    }
}