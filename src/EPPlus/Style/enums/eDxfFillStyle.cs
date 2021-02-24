/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/29/2021         EPPlus Software AB       EPPlus 5.6
 *************************************************************************************************/
namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Fill style for a differential style
    /// </summary>
    public enum eDxfFillStyle
    {
        /// <summary>
        /// Fill using the selected pattern and color
        /// </summary>
        PatternFill,
        /// <summary>
        /// A gradient fill using multiple colors.        
        /// </summary>
        GradientFill
    }
}