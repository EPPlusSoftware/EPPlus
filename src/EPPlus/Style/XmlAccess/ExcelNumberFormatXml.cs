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
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for number formats
    /// </summary>
    public sealed class ExcelNumberFormatXml : StyleXmlHelper
    {
        internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
            
        }        
        internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager, bool buildIn): base(nameSpaceManager)
        {
            BuildIn = buildIn;
        }
        internal ExcelNumberFormatXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _numFmtId = GetXmlNodeInt("@numFmtId");
            _format = GetXmlNodeString("@formatCode");
        }
        /// <summary>
        /// If the number format is build in
        /// </summary>
        public bool BuildIn { get; private set; }
        int _numFmtId;
        /// <summary>
        /// Id for number format
        /// 
        /// Build in ID's
        /// 
        /// 0   General 
        /// 1   0 
        /// 2   0.00 
        /// 3   #,##0 
        /// 4   #,##0.00 
        /// 9   0% 
        /// 10  0.00% 
        /// 11  0.00E+00 
        /// 12  # ?/? 
        /// 13  # ??/?? 
        /// 14  mm-dd-yy 
        /// 15  d-mmm-yy 
        /// 16  d-mmm 
        /// 17  mmm-yy 
        /// 18  h:mm AM/PM 
        /// 19  h:mm:ss AM/PM 
        /// 20  h:mm 
        /// 21  h:mm:ss 
        /// 22  m/d/yy h:mm 
        /// 37  #,##0;(#,##0) 
        /// 38  #,##0;[Red] (#,##0) 
        /// 39  #,##0.00;(#,##0.00) 
        /// 40  #,##0.00;[Red] (#,##0.00) 
        /// 45  mm:ss 
        /// 46  [h]:mm:ss 
        /// 47  mmss.0 
        /// 48  ##0.0E+0 
        /// 49  @
        /// </summary>            
        public int NumFmtId
        {
            get
            {
                return _numFmtId;
            }
            set
            {
                _numFmtId = value;
            }
        }
        internal override string Id
        {
            get
            {
                return _format;
            }
        }
        const string fmtPath = "@formatCode";
        string _format = string.Empty;
        /// <summary>
        /// The numberformat string
        /// </summary>
        public string Format
        {
            get
            {
                return _format;
            }
            set
            {
                _numFmtId = ExcelNumberFormat.GetIdByFormat(value);
                _format = value;
            }
        }
        internal string GetNewID(int NumFmtId, string Format)
        {            
            if (NumFmtId < 0)
            {
                NumFmtId = ExcelNumberFormat.GetIdByEnglishFormat(Format);                
            }
            return NumFmtId.ToString();
        }

        internal static void AddBuildIn(XmlNamespaceManager nameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> numberFormats)
        {
            if (Thread.CurrentThread.CurrentCulture.Name.Equals("de-DE"))
            {
                AddGermanBuildIn(nameSpaceManager, numberFormats);
            } 
            
            AddEnglishBuildIn(nameSpaceManager, numberFormats);
        }

        private static void AddEnglishBuildIn(XmlNamespaceManager nameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> numberFormats)
        {
            numberFormats.Add("General",new ExcelNumberFormatXml(nameSpaceManager,true){NumFmtId=0,Format="General"});
            numberFormats.Add("0", new ExcelNumberFormatXml(nameSpaceManager,true) { NumFmtId = 1, Format = "0" });
            numberFormats.Add("0.00", new ExcelNumberFormatXml(nameSpaceManager,true) { NumFmtId = 2, Format = "0.00" });
            numberFormats.Add("#,##0", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 3, Format = "#,##0" });
            numberFormats.Add("#,##0.00", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 4, Format = "#,##0.00" });
            numberFormats.Add("0%", new ExcelNumberFormatXml(nameSpaceManager,true) { NumFmtId = 9, Format = "0%" });
            numberFormats.Add("0.00%", new ExcelNumberFormatXml(nameSpaceManager,true) { NumFmtId = 10, Format = "0.00%" });
            numberFormats.Add("0.00E+00", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 11, Format = "0.00E+00" });
            numberFormats.Add("# ?/?", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 12, Format = "# ?/?" });
            numberFormats.Add("# ??/??", new ExcelNumberFormatXml(nameSpaceManager,true) { NumFmtId = 13, Format = "# ??/??" });
            numberFormats.Add("mm-dd-yy", new ExcelNumberFormatXml(nameSpaceManager,true) { NumFmtId = 14, Format = "mm-dd-yy" });
            numberFormats.Add("d-mmm-yy", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 15, Format = "d-mmm-yy" });
            numberFormats.Add("d-mmm", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 16, Format = "d-mmm" });
            numberFormats.Add("mmm-yy", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 17, Format = "mmm-yy" });
            numberFormats.Add("h:mm AM/PM", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 18, Format = "h:mm AM/PM" });
            numberFormats.Add("h:mm:ss AM/PM", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 19, Format = "h:mm:ss AM/PM" });
            numberFormats.Add("h:mm", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 20, Format = "h:mm" });
            numberFormats.Add("h:mm:ss", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 21, Format = "h:mm:ss" });
            numberFormats.Add("m/d/yy h:mm", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 22, Format = "m/d/yy h:mm" });
            numberFormats.Add("#,##0 ;(#,##0)", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 37, Format = "#,##0 ;(#,##0)" });
            numberFormats.Add("#,##0 ;[Red](#,##0)", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 38, Format = "#,##0 ;[Red](#,##0)" });
            numberFormats.Add("#,##0.00;(#,##0.00)", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 39, Format = "#,##0.00;(#,##0.00)" });
            numberFormats.Add("#,##0.00;[Red](#,##0.00)", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 40, Format = "#,##0.00;[Red](#,##0.00)" });
            numberFormats.Add("mm:ss", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 45, Format = "mm:ss" });
            numberFormats.Add("[h]:mm:ss", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 46, Format = "[h]:mm:ss" });
            numberFormats.Add("mmss.0", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 47, Format = "mmss.0" });
            numberFormats.Add("##0.0", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 48, Format = "##0.0" });
            numberFormats.Add("@", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 49, Format = "@" });

            numberFormats.NextId = 164; //Start for custom formats.
        }

        private static void AddGermanBuildIn(XmlNamespaceManager nameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> numberFormats)
        {
            numberFormats.Add("General",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 0, Format = "General" });
            numberFormats.Add("0", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 1, Format = "0" });
            numberFormats.Add("0.00",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 2, Format = "0.00" });
            numberFormats.Add("#,##0",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 3, Format = "#,##0" });
            numberFormats.Add("#,##0.00",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 4, Format = "#,##0.00" });
            numberFormats.Add("0%", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 9, Format = "0%" });
            numberFormats.Add("0.00%",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 10, Format = "0.00%" });
            numberFormats.Add("0.00E+00",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 11, Format = "0.00E+00" });
            numberFormats.Add("# ?/?",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 12, Format = "# ?/?" });
            numberFormats.Add("# ??/??",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 13, Format = "# ??/??" });
            numberFormats.Add("dd.mm.yyyy",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 14, Format = "dd.mm.yyyy" });
            numberFormats.Add("dd. mm yy",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 15, Format = "dd. mm yy" });
            numberFormats.Add("dd. mmm",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 16, Format = "dd. mmm" });
            numberFormats.Add("mmm yy",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 17, Format = "mmm yy" });
            numberFormats.Add("h:mm AM/PM",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 18, Format = "h:mm AM/PM" });
            numberFormats.Add("h:mm:ss AM/PM",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 19, Format = "h:mm:ss AM/PM" });
            numberFormats.Add("hh:mm",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 20, Format = "hh:mm" });
            numberFormats.Add("hh:mm:ss",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 21, Format = "hh:mm:ss" });
            numberFormats.Add("dd.mm.yyyy hh:mm",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 22, Format = "dd.mm.yyyy hh:mm" });
            numberFormats.Add("#,##0 _€;-#,##0 _€",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 37, Format = "#,##0 _€;-#,##0 _€" });
            numberFormats.Add("#,##0 _€;[Red]-#,##0 _€",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 38, Format = "#,##0 _€;[Red]-#,##0 _€" });
            numberFormats.Add("#,##0.00 _€;-#,##0.00 _€",
                new ExcelNumberFormatXml(nameSpaceManager, true)
                    { NumFmtId = 39, Format = "#,##0.00 _€;-#,##0.00 _€" });
            numberFormats.Add("#,##0.00 _€;[Red]-#,##0.00 _€",
                new ExcelNumberFormatXml(nameSpaceManager, true)
                    { NumFmtId = 40, Format = "#,##0.00 _€;[Red]-#,##0.00 _€" });
            numberFormats.Add("mm:ss",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 45, Format = "mm:ss" });
            numberFormats.Add("[h]:mm:ss",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 46, Format = "[h]:mm:ss" });
            numberFormats.Add("mm:ss.0",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 47, Format = "mm:ss.0" });
            numberFormats.Add("##0.0E+0",
                new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 48, Format = "##0.0E+0" });
            numberFormats.Add("@", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 49, Format = "@" });

            numberFormats.NextId = 164; //Start for custom formats.
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            SetXmlNodeString("@numFmtId", NumFmtId.ToString());
            SetXmlNodeString("@formatCode", Format);
            return TopNode;
        }

        internal enum eFormatType
        {
            Unknown = 0,
            Number = 1,
            DateTime = 2,
        }
        ExcelFormatTranslator _translator = null;
        internal ExcelFormatTranslator FormatTranslator
        {
            get
            {
                if (_translator == null)
                {
                    _translator = new ExcelFormatTranslator(Format, NumFmtId);
                }
                return _translator;
            }
        }
        #region Excel --> .Net Format
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
                else if (format.Equals("general",StringComparison.OrdinalIgnoreCase))
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
                bool containsAmPm = ExcelFormat.Contains("AM/PM");
                List<int> lstDec=new List<int>();
                StringBuilder sb = new StringBuilder();
                Culture = null;
                char clc;
                var secCount = 0;
                var f = Formats[0];

                if (containsAmPm)
                {
                    ExcelFormat = Regex.Replace(ExcelFormat, "AM/PM", "");
                    DataType = eFormatType.DateTime;
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
                                    string[] li = Regex.Split(bracketText, "-");
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
                                else if(bracketText.StartsWith("<") ||
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
                                    if (c == '.' || c == ',')
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
                                    if(c=='.')
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
                                            sb.Remove(sb.Length-(pos-startPos-1),(pos-startPos-1)) ;

                                        int endPos = pos + 1;
                                        while (endPos < ExcelFormat.Length &&
                                                (ExcelFormat[endPos] == '?' ||
                                                ExcelFormat[endPos] == '#' ||
                                                (ExcelFormat[endPos] >= '0' && ExcelFormat[endPos]<= '9')))
                                        {
                                            endPos++;
                                        }
                                        pos = endPos;
                                        if (f.FractionFormat != "")
                                        {
                                            f.FractionFormat = ExcelFormat.Substring(startPos+1, endPos - startPos-1);
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
                if (numerator == denomerator || numerator==0)
                {
                    if(numerator == denomerator) intPart++;
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
                    for (int i = format.Length - v.Length-1; i >= 0; i--)
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
                if(Formats.Count>1)
                {
                    if(ConvertUtil.IsNumericOrDate(value))
                    {
                        var d=ConvertUtil.GetValueDouble(value);
                        if(d < 0D && Formats.Count > 1)
                        {
                            return Formats[1];
                        }
                        else if(d==0D && Formats.Count > 2)
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
                        if(Formats.Count>3)
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
#endregion
    }
}
