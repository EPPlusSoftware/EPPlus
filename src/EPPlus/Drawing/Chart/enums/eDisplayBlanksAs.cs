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
    /// How to display blanks in a chart
    /// </summary>
    public enum eDisplayBlanksAs
    {
        /// <summary>
        /// Blank values will be left as a gap
        /// </summary>
        Gap,
        /// <summary>
        /// Blank values will be spanned with a line for line charts
        /// </summary>
        Span,
        /// <summary>
        /// Blank values will be treated as zero
        /// </summary>
        Zero
    }
}