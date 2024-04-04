using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using static OfficeOpenXml.Style.XmlAccess.ExcelNumberFormatXml;

namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Translates Excels format to .NET format
    /// </summary>
    internal class ExcelFormatTranslator
    {
        internal enum eSystemDateFormat
        {
            None,
            SystemLongDate,
            SystemLongTime,
            Conditional,
            SystemShortDate,
        }
        internal class FormatPart
        {
            internal string NetFormat { get; set; }
            internal string NetFormatForWidth { get; set; }
            internal string FractionFormat { get; set; }
            internal eSystemDateFormat SpecialDateFormat { get; set; }
            internal bool ContainsTextPlaceholder { get; set; } = false;
            internal void SetFormat(string format, bool containsAmPm, bool forColWidth)
            {
                if (containsAmPm)
                {
                    format += "tt";
                }

                if (forColWidth)
                {
                    NetFormatForWidth = format;
                }
                else
                {
                    NetFormat = format;
                }
            }
        }
        internal ExcelFormatTranslator(string format, int numFmtID)
        {
            var f = new FormatPart();
            Formats.Add(f);
            if (numFmtID == 14)
            {
                f.NetFormat = f.NetFormatForWidth = "";
                DataType = eFormatType.DateTime;
                f.SpecialDateFormat = eSystemDateFormat.SystemShortDate;
            }
            else if (format.Equals("general", StringComparison.OrdinalIgnoreCase))
            {
                f.NetFormat = f.NetFormatForWidth = "0.#########";
                DataType = eFormatType.Number;
            }
            else
            {
                ToNetFormat(format, false);
                ToNetFormat(format, true);
            }
        }

        // escape ('\')  before these characters will be retained
        private static char[] _escapeChars = new char[] { '.', ',', '\'' };

        internal List<FormatPart> Formats { get; private set; } = new List<FormatPart>();
        CultureInfo _ci = null;
        internal CultureInfo Culture
        {
            get
            {
                return _ci ?? CultureInfo.CurrentCulture;
            }
            set
            {
                _ci = value;
            }
        }
        internal bool HasCulture
        {
            get
            {
                return _ci != null;
            }
        }
        internal eFormatType DataType { get; private set; }
        private void ToNetFormat(string ExcelFormat, bool forColWidth)
        {
            DataType = eFormatType.Unknown;
            bool isText = false;
            bool isBracket = false;
            string bracketText = "";
            bool prevBslsh = false;
            bool useMinute = false;
            bool prevUnderScore = false;
            bool ignoreNext = false;
            bool containsAmPm = ExcelFormat.IndexOf("AM/PM", StringComparison.InvariantCultureIgnoreCase) >= 0;
            List<int> lstDec = new List<int>();
            StringBuilder sb = new StringBuilder();
            Culture = null;
            char clc;
            var secCount = 0;
            var f = Formats[0];

            if (containsAmPm)
            {
                ExcelFormat = Regex.Replace(ExcelFormat, "AM/PM", "", RegexOptions.IgnoreCase);
            }

            for (int pos = 0; pos < ExcelFormat.Length; pos++)
            {
                char c = ExcelFormat[pos];
                if (c == '"')
                {
                    isText = !isText;
                    sb.Append(c);
                }
                else
                {
                    if (ignoreNext)
                    {
                        ignoreNext = false;
                        continue;
                    }
                    else if (isText && !isBracket)
                    {
                        sb.Append(c);
                    }
                    else if (isBracket)
                    {
                        if (c == ']')
                        {
                            isBracket = false;
                            if (bracketText[0] == '$')  //Local Info
                            {
                                //string[] li = Regex.Split(bracketText, "-");
                                string[] li = bracketText.Split('-');
                                if (li[0].Length > 1)
                                {
                                    sb.Append("\"" + li[0].Substring(1, li[0].Length - 1) + "\"");     //Currency symbol
                                }
                                if (li.Length > 1)
                                {
                                    if (li[1].Equals("f800", StringComparison.OrdinalIgnoreCase))
                                    {
                                        f.SpecialDateFormat = eSystemDateFormat.SystemLongDate;
                                    }
                                    else if (li[1].Equals("f400", StringComparison.OrdinalIgnoreCase))
                                    {
                                        f.SpecialDateFormat = eSystemDateFormat.SystemLongTime;
                                    }
                                    else if (int.TryParse(li[1], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int num))
                                    {
                                        try
                                        {
                                            Culture = CultureInfo.GetCultureInfo(num & 0xFFFF);
                                        }
                                        catch
                                        {
                                            Culture = null;
                                        }
                                    }
                                    else //Excel saves in hex, but seems to support having culture codes as well.
                                    {
                                        try
                                        {
                                            Culture = CultureInfo.GetCultureInfo(li[1]);
                                        }
                                        catch
                                        {
                                            Culture = null;
                                        }
                                    }
                                }
                            }
                            else if (bracketText.StartsWith("<") ||
                                    bracketText.StartsWith(">") ||
                                    bracketText.StartsWith("=")) //Conditional
                            {
                                f.SpecialDateFormat = eSystemDateFormat.Conditional;
                            }
                            else
                            {
                                sb.Append(bracketText);
                                f.SpecialDateFormat = eSystemDateFormat.Conditional;
                            }
                        }
                        else
                        {
                            bracketText += c;
                        }
                    }
                    else if (prevUnderScore)
                    {
                        if (forColWidth)
                        {
                            sb.AppendFormat("\"{0}\"", c);
                        }
                        prevUnderScore = false;
                    }
                    else
                    {
                        if (c == ';') //We use first part (for positive only at this stage)
                        {
                            secCount++;
                            f.SetFormat(sb.ToString(), containsAmPm, forColWidth);
                            if (secCount < Formats.Count)
                            {
                                f = Formats[secCount];
                            }
                            else
                            {
                                f = new FormatPart();
                                Formats.Add(f);
                            }
                            sb = new StringBuilder();
                        }
                        else
                        {
                            clc = c.ToString().ToLower(CultureInfo.InvariantCulture)[0];  //Lowercase character
                                                                                          //Set the datetype
                            if (DataType == eFormatType.Unknown)
                            {
                                if (c == '0' || c == '#' || c == '.')
                                {
                                    DataType = eFormatType.Number;
                                }
                                else if (clc == 'y' || clc == 'm' || clc == 'd' || clc == 'h' || clc == 'm' || clc == 's')
                                {
                                    DataType = eFormatType.DateTime;
                                }
                            }

                            if (prevBslsh)
                            {
                                if (_escapeChars.Contains(c))
                                {
                                    sb.Append('\\');
                                }
                                sb.Append(c);
                                prevBslsh = false;
                            }
                            else if (c == '[')
                            {
                                bracketText = "";
                                isBracket = true;
                            }
                            else if (c == '\\')
                            {
                                prevBslsh = true;
                            }
                            else if (c == '0' ||
                                c == '#' ||
                                c == '.' ||
                                c == ',' ||
                                c == '%' ||
                                clc == 'd' ||
                                clc == 's')
                            {
                                sb.Append(c);
                                if (c == '.')
                                {
                                    lstDec.Add(sb.Length - 1);
                                }
                            }
                            else if (clc == 'h')
                            {
                                if (containsAmPm)
                                {
                                    sb.Append('h');
                                }
                                else
                                {
                                    sb.Append('H');
                                }
                                useMinute = true;
                            }
                            else if (clc == 'm')
                            {
                                if (useMinute)
                                {
                                    sb.Append('m');
                                }
                                else
                                {
                                    sb.Append('M');
                                }
                            }
                            else if (c == '_') //Skip next but use for alignment
                            {
                                prevUnderScore = true;
                            }
                            else if (c == '?')
                            {
                                sb.Append(' ');
                            }
                            else if (c == '/')
                            {
                                if (DataType == eFormatType.Number)
                                {
                                    int startPos = pos - 1;
                                    while (startPos >= 0 &&
                                            (ExcelFormat[startPos] == '?' ||
                                            ExcelFormat[startPos] == '#' ||
                                            ExcelFormat[startPos] == '0'))
                                    {
                                        startPos--;
                                    }

                                    if (startPos > 0)  //RemovePart
                                        sb.Remove(sb.Length - (pos - startPos - 1), (pos - startPos - 1));

                                    int endPos = pos + 1;
                                    while (endPos < ExcelFormat.Length &&
                                            (ExcelFormat[endPos] == '?' ||
                                            ExcelFormat[endPos] == '#' ||
                                            (ExcelFormat[endPos] >= '0' && ExcelFormat[endPos] <= '9')))
                                    {
                                        endPos++;
                                    }
                                    pos = endPos;
                                    if (f.FractionFormat != "")
                                    {
                                        f.FractionFormat = ExcelFormat.Substring(startPos + 1, endPos - startPos - 1);
                                    }
                                    sb.Append('?'); //Will be replaced later on by the fraction
                                }
                                else
                                {
                                    sb.Append('/');
                                }
                            }
                            else if (c == '*')
                            {
                                //repeat char--> ignore
                                ignoreNext = true;
                            }
                            else if (c == '@')
                            {
                                sb.Append("{0}");
                                f.ContainsTextPlaceholder = true;
                            }
                            else
                            {
                                sb.Append(c);
                            }
                        }
                    }
                }
            }

            //Add qoutes
            if (DataType == eFormatType.DateTime) SetDecimal(lstDec, sb); //Remove?


            //if (format == "")
            //    format = sb.ToString();
            //else
            //    text = sb.ToString();

            // AM/PM format
            f.SetFormat(sb.ToString(), containsAmPm, forColWidth);
        }

        private static void SetDecimal(List<int> lstDec, StringBuilder sb)
        {
            if (lstDec.Count > 1)
            {
                for (int i = lstDec.Count - 1; i >= 0; i--)
                {
                    sb.Insert(lstDec[i] + 1, '\'');
                    sb.Insert(lstDec[i], '\'');
                }
            }
        }

        internal string FormatFraction(double d, FormatPart f)
        {
            int numerator, denomerator;

            int intPart = (int)d;

            string[] fmt = f.FractionFormat.Split('/');

            int fixedDenominator;
            if (!int.TryParse(fmt[1], out fixedDenominator))
            {
                fixedDenominator = 0;
            }

            if (d == 0 || double.IsNaN(d))
            {
                if (fmt[0].Trim() == "" && fmt[1].Trim() == "")
                {
                    return new string(' ', f.FractionFormat.Length);
                }
                else
                {
                    return 0.ToString(fmt[0]) + "/" + 1.ToString(fmt[0]);
                }
            }

            int maxDigits = fmt[1].Length;
            string sign = d < 0 ? "-" : "";
            if (fixedDenominator == 0)
            {
                List<double> numerators = new List<double>() { 1, 0 };
                List<double> denominators = new List<double>() { 0, 1 };

                if (maxDigits < 1 && maxDigits > 12)
                {
                    throw (new ArgumentException("Number of digits out of range (1-12)"));
                }

                int maxNum = 0;
                for (int i = 0; i < maxDigits; i++)
                {
                    maxNum += 9 * (int)(Math.Pow((double)10, (double)i));
                }

                double divRes = 1 / ((double)Math.Abs(d) - intPart);
                double result, prevResult = double.NaN;
                int listPos = 2, index = 1;
                while (true)
                {
                    index++;
                    double intDivRes = Math.Floor(divRes);
                    numerators.Add((intDivRes * numerators[index - 1] + numerators[index - 2]));
                    if (numerators[index] > maxNum)
                    {
                        break;
                    }

                    denominators.Add((intDivRes * denominators[index - 1] + denominators[index - 2]));

                    result = numerators[index] / denominators[index];
                    if (denominators[index] > maxNum)
                    {
                        break;
                    }
                    listPos = index;

                    if (result == prevResult) break;

                    if (result == d) break;

                    prevResult = result;

                    divRes = 1 / (divRes - intDivRes);  //Rest
                }

                numerator = (int)numerators[listPos];
                denomerator = (int)denominators[listPos];
            }
            else
            {
                numerator = (int)Math.Round((d - intPart) / (1D / fixedDenominator), 0);
                denomerator = fixedDenominator;
            }
            if (numerator == denomerator || numerator == 0)
            {
                if (numerator == denomerator) intPart++;
                return sign + intPart.ToString(f.NetFormat).Replace("?", new string(' ', f.FractionFormat.Length));
            }
            else if (intPart == 0)
            {
                return sign + FmtInt(numerator, fmt[0]) + "/" + FmtInt(denomerator, fmt[1]);
            }
            else
            {
                return sign + intPart.ToString(f.NetFormat).Replace("?", FmtInt(numerator, fmt[0]) + "/" + FmtInt(denomerator, fmt[1]));
            }
        }

        private string FmtInt(double value, string format)
        {
            string v = value.ToString("#");
            string pad = "";
            if (v.Length < format.Length)
            {
                for (int i = format.Length - v.Length - 1; i >= 0; i--)
                {
                    if (format[i] == '?')
                    {
                        pad += " ";
                    }
                    else if (format[i] == ' ')
                    {
                        pad += "0";
                    }
                }
            }
            return pad + v;
        }

        internal FormatPart GetFormatPart(object value)
        {
            if (Formats.Count > 1)
            {
                if (ConvertUtil.IsNumericOrDate(value))
                {
                    var d = ConvertUtil.GetValueDouble(value);
                    if (d < 0D && Formats.Count > 1)
                    {
                        return Formats[1];
                    }
                    else if (d == 0D && Formats.Count > 2)
                    {
                        return Formats[2];
                    }
                    else
                    {
                        return Formats[0];
                    }
                }
                else
                {
                    if (Formats.Count > 3)
                    {
                        return Formats[3];
                    }
                    else
                    {
                        return Formats[0];
                    }
                }
            }
            else
            {
                return Formats[0];
            }
        }
    }
}
