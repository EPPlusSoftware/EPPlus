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
using System.Threading;

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

        internal static void AddBuiltIn(
            XmlNamespaceManager nameSpaceManager,
            ExcelStyleCollection<ExcelNumberFormatXml> numberFormats)
        {
            AddDefinableFormats(nameSpaceManager, numberFormats);

            if (Thread.CurrentThread.CurrentCulture.Name.Equals("de-DE"))
            {
                AddNonDefinableGermanFormats(nameSpaceManager, numberFormats);
            }
            else
            {
                AddDefaultFormats(nameSpaceManager, numberFormats);
            }
        }

        private static void AddDefinableFormats(XmlNamespaceManager nameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> numberFormats)
        {
            numberFormats.Add("General", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 0, Format = "General" });
            numberFormats.Add("0", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 1, Format = "0" });
            numberFormats.Add("0.00", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 2, Format = "0.00" });
            numberFormats.Add("#,##0", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 3, Format = "#,##0" });
            numberFormats.Add("#,##0.00", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 4, Format = "#,##0.00" });
            numberFormats.Add("0%", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 9, Format = "0%" });
            numberFormats.Add("0.00%", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 10, Format = "0.00%" });
            numberFormats.Add("0.00E+00", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 11, Format = "0.00E+00" });
            numberFormats.Add("# ?/?", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 12, Format = "# ?/?" });
            numberFormats.Add("# ??/??", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 13, Format = "# ??/??" });
            numberFormats.Add("h:mm AM/PM", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 18, Format = "h:mm AM/PM" });
            numberFormats.Add("h:mm:ss AM/PM", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 19, Format = "h:mm:ss AM/PM" });
            numberFormats.Add("mm:ss", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 45, Format = "mm:ss" });
            numberFormats.Add("[h]:mm:ss", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 46, Format = "[h]:mm:ss" });
            numberFormats.Add("@", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 49, Format = "@" });

            numberFormats.NextId = 164; // Start for custom formats.
        }

        private static void AddNonDefinableGermanFormats(XmlNamespaceManager nameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> numberFormats)
        {
            numberFormats.Add("dd.mm.yyyy", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 14, Format = "dd.mm.yyyy" });
            numberFormats.Add("dd. mm yy", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 15, Format = "dd. mm yy" });
            numberFormats.Add("dd. mmm", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 16, Format = "dd. mmm" });
            numberFormats.Add("mmm yy", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 17, Format = "mmm yy" });
            numberFormats.Add("hh:mm", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 20, Format = "hh:mm" });
            numberFormats.Add("hh:mm:ss", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 21, Format = "hh:mm:ss" });
            numberFormats.Add("dd.mm.yyyy hh:mm", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 22, Format = "dd.mm.yyyy hh:mm" });
            numberFormats.Add("#,##0 _€;-#,##0 _€", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 37, Format = "#,##0 _€;-#,##0 _€" });
            numberFormats.Add("#,##0 _€;[Red]-#,##0 _€", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 38, Format = "#,##0 _€;[Red]-#,##0 _€" });
            numberFormats.Add("#,##0.00 _€;-#,##0.00 _€", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 39, Format = "#,##0.00 _€;-#,##0.00 _€" });
            numberFormats.Add("#,##0.00 _€;[Red]-#,##0.00 _€", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 40, Format = "#,##0.00 _€;[Red]-#,##0.00 _€" });
            numberFormats.Add("mm:ss.0", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 47, Format = "mm:ss.0" });
            numberFormats.Add("##0.0E+0", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 48, Format = "##0.0E+0" });
        }

        private static void AddDefaultFormats(XmlNamespaceManager nameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> numberFormats)
        {
            numberFormats.Add("mm-dd-yy", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 14, Format = "mm-dd-yy" });
            numberFormats.Add("d-mmm-yy", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 15, Format = "d-mmm-yy" });
            numberFormats.Add("d-mmm", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 16, Format = "d-mmm" });
            numberFormats.Add("mmm-yy", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 17, Format = "mmm-yy" });
            numberFormats.Add("h:mm", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 20, Format = "h:mm" });
            numberFormats.Add("h:mm:ss", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 21, Format = "h:mm:ss" });
            numberFormats.Add("m/d/yy h:mm", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 22, Format = "m/d/yy h:mm" });
            numberFormats.Add("#,##0 ;(#,##0)", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 37, Format = "#,##0 ;(#,##0)" });
            numberFormats.Add("#,##0 ;[Red](#,##0)", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 38, Format = "#,##0 ;[Red](#,##0)" });
            numberFormats.Add("#,##0.00;(#,##0.00)", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 39, Format = "#,##0.00;(#,##0.00)" });
            numberFormats.Add("#,##0.00;[Red](#,##0.00)", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 40, Format = "#,##0.00;[Red](#,##0.00)" });
            numberFormats.Add("mmss.0", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 47, Format = "mmss.0" });
            numberFormats.Add("##0.0", new ExcelNumberFormatXml(nameSpaceManager, true) { NumFmtId = 48, Format = "##0.0" });
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
