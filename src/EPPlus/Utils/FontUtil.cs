

using OfficeOpenXml.Interfaces.Drawing.Text;

namespace OfficeOpenXml.Utils
{
    internal class FontUtil
    {
        internal static MeasurementFont GetMeasureFont(string latinFont, string complexFont, string eastAsianFont, float size, MeasurementFontStyles style, ExcelPackage pck)
        {
            var mf = new MeasurementFont()
            {
                FontFamily = string.IsNullOrEmpty(latinFont) ? complexFont : latinFont,
                Size = size,
                Style = style
            };
            if (string.IsNullOrEmpty(mf.FontFamily))
            {
                mf.FontFamily = eastAsianFont;
            }

            switch (mf.FontFamily)
            {
                case "+mn-lt":
                    mf.FontFamily = pck.Workbook.ThemeManager.GetOrCreateTheme().FontScheme.MinorFont.LatinFont?.Typeface;
                    break;
                case "+mj-lt":
                    mf.FontFamily = pck.Workbook.ThemeManager.GetOrCreateTheme().FontScheme.MajorFont.LatinFont?.Typeface;
                    break;
                case "+mn-cs":
                    mf.FontFamily = pck.Workbook.ThemeManager.GetOrCreateTheme().FontScheme.MinorFont.ComplexFont?.Typeface;
                    break;
                case "+mj-cs":
                    mf.FontFamily = pck.Workbook.ThemeManager.GetOrCreateTheme().FontScheme.MajorFont.ComplexFont?.Typeface;
                    break;
                case "+mn-ea":
                    mf.FontFamily = pck.Workbook.ThemeManager.GetOrCreateTheme().FontScheme.MinorFont.EastAsianFont?.Typeface;
                    break;
                case "+mj-ea":
                    mf.FontFamily = pck.Workbook.ThemeManager.GetOrCreateTheme().FontScheme.MajorFont.EastAsianFont?.Typeface;
                    break;
            }

            if (string.IsNullOrEmpty(mf.FontFamily) || mf.Size <= 0 || double.IsNaN(mf.Size))
            {
                var ns = pck.Workbook.Styles.GetNormalStyle();
                if (string.IsNullOrEmpty(mf.FontFamily))
                {
                    if (ns == null || string.IsNullOrEmpty(ns.Style.Font.Name))
                    {
                        mf.FontFamily = "Aptos Narrow";
                    }
                    else
                    {
                        mf.FontFamily = ns.Style.Font.Name;
                    }
                }

                if (mf.Size <= 0 || double.IsNaN(mf.Size))
                {
                    mf.Size = ns.Style.Font.Size;
                }
            }

            return mf;

        }
    }
}
