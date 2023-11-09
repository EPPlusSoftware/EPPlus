using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    /// <summary>
    /// 
    /// </summary>
    public interface IFill
    {
        double Degree { get; }
        double Right { get; }
        double Bottom { get; }
        bool IsLinear { get; }
        bool IsGradient { get; }

        bool HasValue { get; }
        //string Color1 { get; }
        //string Color2 { get; }

        /// <summary>
        /// 
        /// </summary>
        ExcelFillStyle PatternType { get; }

        string GetBackgroundColor(ExcelTheme theme);
        string GetPatternColor(ExcelTheme theme);
        string GetGradientColor1(ExcelTheme theme);
        string GetGradientColor2(ExcelTheme theme);
    }
}
