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
    /// Text anchoring
    /// </summary>
    public enum eTextAnchoringType  
    {
        /// <summary>
        /// Anchor the text to the bottom
        /// </summary>
        Bottom,
        /// <summary>
        /// Anchor the text to the center
        /// </summary>
        Center,
        /// <summary>
        /// Anchor the text so that it is distributed vertically.
        /// </summary>
        Distributed,
        /// <summary>
        /// Anchor the text so that it is justified vertically.
        /// </summary>
        Justify,
        /// <summary>
        /// Anchor the text to the top
        /// </summary>
        Top
    }
}