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
            NumberFormats.Add("General",new ExcelNumberFormatXml(NameSpaceManager,true){NumFmtId=0,Format="General"});
            NumberFormats.Add("0", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 1, Format = "0" });
            NumberFormats.Add("0.00", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 2, Format = "0.00" });
            NumberFormats.Add("#,##0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 3, Format = "#,##0" });
            NumberFormats.Add("#,##0.00", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 4, Format = "#,##0.00" });
            NumberFormats.Add("0%", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 9, Format = "0%" });
            NumberFormats.Add("0.00%", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 10, Format = "0.00%" });
            NumberFormats.Add("0.00E+00", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 11, Format = "0.00E+00" });
            NumberFormats.Add("# ?/?", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 12, Format = "# ?/?" });
            NumberFormats.Add("# ??/??", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 13, Format = "# ??/??" });
            NumberFormats.Add("mm-dd-yy", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 14, Format = "mm-dd-yy" });
            NumberFormats.Add("d-mmm-yy", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 15, Format = "d-mmm-yy" });
            NumberFormats.Add("d-mmm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 16, Format = "d-mmm" });
            NumberFormats.Add("mmm-yy", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 17, Format = "mmm-yy" });
            NumberFormats.Add("h:mm AM/PM", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 18, Format = "h:mm AM/PM" });
            NumberFormats.Add("h:mm:ss AM/PM", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 19, Format = "h:mm:ss AM/PM" });
            NumberFormats.Add("h:mm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 20, Format = "h:mm" });
            NumberFormats.Add("h:mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 21, Format = "h:mm:ss" });
            NumberFormats.Add("m/d/yy h:mm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 22, Format = "m/d/yy h:mm" });
            NumberFormats.Add("#,##0 ;(#,##0)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 37, Format = "#,##0 ;(#,##0)" });
            NumberFormats.Add("#,##0 ;[Red](#,##0)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 38, Format = "#,##0 ;[Red](#,##0)" });
            NumberFormats.Add("#,##0.00;(#,##0.00)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 39, Format = "#,##0.00;(#,##0.00)" });
            NumberFormats.Add("#,##0.00;[Red](#,##0.00)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 40, Format = "#,##0.00;[Red](#,##0.00)" });
            NumberFormats.Add("mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 45, Format = "mm:ss" });
            NumberFormats.Add("[h]:mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 46, Format = "[h]:mm:ss" });
            NumberFormats.Add("mmss.0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 47, Format = "mmss.0" });
            NumberFormats.Add("##0.0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 48, Format = "##0.0" });
            NumberFormats.Add("@", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 49, Format = "@" });

            NumberFormats.NextId = 164; //Start for custom formats.
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
