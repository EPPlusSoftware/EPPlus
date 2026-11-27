using OfficeOpenXml.Compatibility;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeOpenXml.Utils.String
{
    /// <summary>
    /// Creates display strings for a value, widest first, using the same formatting as ExcelRangeBase.Text.
    /// The caller measures each candidate and uses the first one that fits.
    /// Only General-formatted numbers get shorter candidates. Explicit formats like "0.000" are
    /// returned as-is, since Excel doesn't reduce decimals for them.
    /// </summary>
    internal static class NumberFitFormatter
    {
        // ExcelFormatTranslator formats General as G10.
        private const int GeneralSignificantDigits = 10;
        // Excel General shows at most 5 mantissa decimals in scientific notation, e.g. 1.23457E+11
        private const int MaxScientificDecimals = 5;

        private static readonly ConcurrentDictionary<int, ExcelFormatTranslator> _scientificFormats =
            new ConcurrentDictionary<int, ExcelFormatTranslator>();

        internal static IEnumerable<string> GetCandidates(object value, ExcelFormatTranslator nf, CultureInfo culture = null)
        {
            // 1. Exactly what ExcelRangeBase.Text returns.
            var text = ValueToTextHandler.FormatValue(value, false, nf, culture, out bool isValid);
            if (text != null) yield return text;

            if (isValid == false || IsGeneralNumber(value, nf) == false) yield break;

            var d = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (d == 0D || double.IsNaN(d) || double.IsInfinity(d)) yield break;

            var prev = text;
            var isScientific = text != null && text.IndexOf('E') >= 0;

            // 2. Fixed notation, one decimal fewer each step. Keep the notation .Text already chose.
            if (isScientific == false)
            {
                var intDigits = Math.Max(1, (int)RoundingHelper.GetNumberOfDigitsIntPart(Math.Abs(d)));
                for (var dec = Math.Max(0, GeneralSignificantDigits - intDigits); dec >= 0; dec--)
                {
                    var rounded = RoundingHelper.RoundToDecimals(d, dec);
                    var s = ValueToTextHandler.FormatValue(rounded, false, nf, culture, out _);
                    if (s != null && s != prev)
                    {
                        yield return prev = s;
                    }
                    if (rounded == 0D) yield break;   // e.g. 0.4 -> "0". Nothing shorter is meaningful.
                }
            }

            // 3. Scientific notation, used when the integer part itself doesn't fit.
            for (var dec = MaxScientificDecimals; dec >= 0; dec--)
            {
                var s = ValueToTextHandler.FormatValue(d, false, GetScientificFormat(dec), culture ?? nf.Culture, out _);
                if (s != null && s != prev)
                {
                    yield return prev = s;
                }
            }
        }

        private static bool IsGeneralNumber(object value, ExcelFormatTranslator nf)
        {
            if (value is bool) return false;
            if ((value is decimal || TypeCompat.IsPrimitive(value)) == false) return false;   // Same test as FormatValue
            return nf.DataType == ExcelNumberFormatXml.eFormatType.Number
                && nf.Formats.Count == 1
                && nf.Formats[0].SpecialDateFormat == ExcelFormatTranslator.eSystemDateFormat.General;
        }

        private static ExcelFormatTranslator GetScientificFormat(int decimals)
        {
            return _scientificFormats.GetOrAdd(decimals, dec =>
                new ExcelFormatTranslator((dec == 0 ? "0" : "0." + new string('#', dec)) + "E+00", -1));
        }
    }
}