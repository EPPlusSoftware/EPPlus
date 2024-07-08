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
using OfficeOpenXml.Utils;
using System.Runtime.InteropServices;

namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for number customFormats
    /// </summary>
    public sealed class ExcelNumberFormatXml : StyleXmlHelper
    {

        internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {

        }
        internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager, bool buildIn) : base(nameSpaceManager)
        {
            BuildIn = buildIn;
        }
        internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager, bool buildIn, int numFmtId, string format) : base(nameSpaceManager)
        {
            BuildIn = buildIn;
            _numFmtId = numFmtId;
            _format = format;
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
                _numFmtId = ExcelNumberFormat.GetFromBuildIdFromFormat(value);
                _format = value;
            }
        }
        internal string GetNewID(int NumFmtId, string Format)
        {
            if (NumFmtId < 0)
            {
                NumFmtId = ExcelNumberFormat.GetFromBuildIdFromFormat(Format);
            }
            return NumFmtId.ToString();
        }

        internal static void AddBuildIn(XmlNamespaceManager NameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> NumberFormats)
        {
            var customFormats = ExcelPackageSettings.CultureSpecificBuildInNumberFormats.ContainsKey(CultureInfo.CurrentCulture.Name) ? ExcelPackageSettings.CultureSpecificBuildInNumberFormats[CultureInfo.CurrentCulture.Name] : null;

            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 0, "General");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 1, "0");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 2,"0.00");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 3, "#,##0");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 4, "#,##0.00");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 9, "0%");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 10, "0.00%");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 11, "0.00E+00");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 12, "# ?/?");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 13, "# ??/??");

            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 14, "mm-dd-yy");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 15, "d-mmm-yy");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 16, "d-mmm");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 17, "mmm-yy");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 18, "h:mm AM/PM");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 19, "h:mm:ss AM/PM");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 20, "h:mm");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 21, "h:mm:ss");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 22, "m/d/yy h:mm");

            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 37, "#,##0 ;(#,##0)");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 38, "#,##0 ;[Red](#,##0)");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 39, "#,##0.00;(#,##0.00)");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 40, "#,##0.00;[Red](#,##0.00)");

            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 45, "mm:ss");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 46, "[h]:mm:ss");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 47, "mmss.0");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 48,"##0.0E+0");
            AddLocalizedFormat(NameSpaceManager, NumberFormats, customFormats, 49, "@");

            NumberFormats.NextId = 164; //Start for custom customFormats.
        }

        private static void AddLocalizedFormat(XmlNamespaceManager nameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> numberFormats, Dictionary<int, string> customFormats, int numFmtId, string format)
        {
            if (customFormats != null && customFormats.TryGetValue(numFmtId, out string customFormat))
            {
                numberFormats.Add(customFormat, new ExcelNumberFormatXml(nameSpaceManager, true, numFmtId, customFormat));
            }
            else
            {
                numberFormats.Add(format, new ExcelNumberFormatXml(nameSpaceManager, true, numFmtId, format));
            }
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
    }
}
