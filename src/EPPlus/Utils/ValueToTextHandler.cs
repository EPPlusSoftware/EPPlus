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
using OfficeOpenXml;
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal static class ValueToTextHandler
    {
        internal static string GetFormattedText(object Value, ExcelWorkbook wb, int styleId, bool forWidthCalc, CultureInfo cultureInfo=null)
        {
            object v = Value;
            if (v == null) return "";
            var styles = wb.Styles;
            var nfID = styles.CellXfs[styleId].NumberFormatId;
            ExcelNumberFormatXml.ExcelFormatTranslator nf = null;
            for (int i = 0; i < styles.NumberFormats.Count; i++)
            {
                if (nfID == styles.NumberFormats[i].NumFmtId)
                {
                    nf = styles.NumberFormats[i].FormatTranslator;
                    break;
                }
            }
            if (nf == null)
            {
                nf = styles.NumberFormats[0].FormatTranslator;  //nf should never be null. If so set to General, Issue 173
            }

            return FormatValue(v, forWidthCalc, nf, cultureInfo);
        }
        internal static string FormatValue(object v, bool forWidthCalc, ExcelNumberFormatXml.ExcelFormatTranslator nf, CultureInfo overrideCultureInfo)
        {
            var f = nf.GetFormatPart(v);
            string format;
            if (forWidthCalc)
            {
                format = f.NetFormatForWidth;
            }
            else
            {
                format = f.NetFormat;
            }


            if (v is decimal || TypeCompat.IsPrimitive(v))
            {
                double d;
                try
                {
                    d = Convert.ToDouble(v);
                }
                catch
                {
                    return "";
                }

                if (nf.DataType == ExcelNumberFormatXml.eFormatType.Number)
                {
                    if (string.IsNullOrEmpty(f.FractionFormat))
                    {                        
                        return FormatNumber(d, format, overrideCultureInfo ?? nf.Culture);
                    }
                    else
                    {
                        return nf.FormatFraction(d, f);
                    }
                }
                else if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
                {
                    if (d > 0)
                    {
                        var date = DateTime.FromOADate(d);
                        return GetDateText(date, format, f, overrideCultureInfo ?? nf.Culture);
                    }
                }

                if (nf.Formats.Count > 2 && string.IsNullOrEmpty(f.NetFormat))
                {
                    return null;
                }
                else if (string.IsNullOrEmpty(format)==false)
                {
                    return d.ToString(format);
                }
            }
            else if (v is DateTime dt)
            {
                if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
                {
                    return GetDateText(dt, format, f, overrideCultureInfo ?? nf.Culture);
                }
                else
                {
                    double d = (dt).ToOADate();
                    if (string.IsNullOrEmpty(f.FractionFormat))
                    {
                        return d.ToString(format, nf.Culture);
                    }
                    else
                    {
                        return nf.FormatFraction(d, f);
                    }
                }
            }
            else if (v is TimeSpan ts)
            {
                if (nf.DataType == ExcelNumberFormatXml.eFormatType.DateTime)
                {
                    return GetDateText(new DateTime(ts.Ticks), format,f, overrideCultureInfo);
                }
                else
                {
                    double d = new DateTime(0).Add(ts).ToOADate();
                    if (string.IsNullOrEmpty(f.FractionFormat))
                    {
                        return d.ToString(format, nf.Culture);
                    }
                    else
                    {
                        return nf.FormatFraction(d,f);
                    }
                }                
            }
            else
            {
                if (nf.Formats.Count > 3 && string.IsNullOrEmpty(f.NetFormat))
                {
                    return null;
                }

                if (f.ContainsTextPlaceholder)
                {
                    return string.Format(format.Replace("\"",""), v);
                }
                else
                {
                    return v.ToString();
                }
            }

            return v.ToString();
        }

        private static string FormatNumber(double d, string format, CultureInfo cultureInfo)
        {
            var s = FormatNumberExcel(d, format, cultureInfo);
            if (string.IsNullOrEmpty(s) == false && (
                    s.StartsWith("--") && format.StartsWith("-") ||
                   (s.StartsWith("-(", StringComparison.OrdinalIgnoreCase) && format.StartsWith("(", StringComparison.OrdinalIgnoreCase) && format.IndexOf(")", StringComparison.OrdinalIgnoreCase)>0)))
            {
                return s.Substring(1);
            }
            else
            {
                return s;
            }
        }

        private static string FormatNumberExcel(double d, string format, CultureInfo cultureInfo)
        {
            if (string.IsNullOrEmpty(format))
            {
                return null;
            }
            else
            {
                return d.ToString(format, cultureInfo);
            }
        }

        private static string GetDateText(DateTime d, string format, ExcelNumberFormatXml.ExcelFormatTranslator.FormatPart f, CultureInfo cultureInfo)
        {           
            if (f.SpecialDateFormat == ExcelNumberFormatXml.ExcelFormatTranslator.eSystemDateFormat.SystemLongDate)
            {
                return d.ToLongDateString();
            }
            else if (f.SpecialDateFormat == ExcelNumberFormatXml.ExcelFormatTranslator.eSystemDateFormat.SystemLongTime)
            {
                return d.ToLongTimeString();
            }
            else if (f.SpecialDateFormat == ExcelNumberFormatXml.ExcelFormatTranslator.eSystemDateFormat.SystemShortDate)
            {
                return d.ToShortDateString();
            }
            if (format == "d" || format == "D")
            {
                return d.Day.ToString();
            }
            else if (format == "M")
            {
                return d.Month.ToString();
            }
            else if (format == "m")
            {
                return d.Minute.ToString();
            }
            else if (format.ToLower() == "y" || format.ToLower() == "yy")
            {
                return d.ToString("yy", cultureInfo);
            }
            else if (format.ToLower() == "yyy" || format.ToLower() == "yyyy")
            {
                return d.ToString("yyy", cultureInfo);
            }
            else
            {
                return d.ToString(format, cultureInfo);
            }

        }
    }
}
