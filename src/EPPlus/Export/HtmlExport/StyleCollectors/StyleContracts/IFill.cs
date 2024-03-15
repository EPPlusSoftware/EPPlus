/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    /// <summary>
    /// For internal use
    /// </summary>
    public interface IFill
    {
        internal double Degree { get; }
        internal double Right { get; }
        internal double Bottom { get; }
        internal bool IsLinear { get; }
        internal bool IsGradient { get; }

        internal bool HasValue { get; }

        /// <summary>
        /// 
        /// </summary>
        internal ExcelFillStyle PatternType { get; }

        internal string GetBackgroundColor(ExcelTheme theme);
        internal string GetPatternColor(ExcelTheme theme);
        internal string GetGradientColor1(ExcelTheme theme);
        internal string GetGradientColor2(ExcelTheme theme);
    }
}
