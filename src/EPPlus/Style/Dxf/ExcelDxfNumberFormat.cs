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
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style.Dxf
{
    /// <summary>
    /// A numberformat in a differential formatting record
    /// </summary>
    public class ExcelDxfNumberFormat : DxfStyleBase
    {
        internal ExcelDxfNumberFormat(ExcelStyles styles, Action<eStyleClass, eStyleProperty, object> callback) : base(styles, callback)
        {

        }
        int _numFmtID=int.MinValue;
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
        /// 37  #,##0 ;(#,##0) 
        /// 38  #,##0 ;[Red](#,##0) 
        /// 39  #,##0.00;(#,##0.00) 
        /// 40  #,##0.00;[Red](#,##0.00) 
        /// 45  mm:ss 
        /// 46  [h]:mm:ss 
        /// 47  mmss.0 
        /// 48  ##0.0E+0 
        /// 49  @
        /// </summary>            
        public int NumFmtID 
        { 
            get
            {
                return _numFmtID;
            }
            internal set
            {
                _numFmtID = value;
            }
        }
        string _format="";
        /// <summary>
        /// The number format
        /// </summary>s
        public string Format
        { 
            get
            {
                return _format;
            }
            set
            {
                _format = value;
                NumFmtID = ExcelNumberFormat.GetFromBuildIdFromFormat(value);
                _callback?.Invoke(eStyleClass.Numberformat, eStyleProperty.Format, value);
            }
        }

        /// <summary>
        /// The id
        /// </summary>
        internal override string Id
        {
            get
            {
                return Format;
            }
        }

        /// <summary>
        /// Creates the the xml node
        /// </summary>
        /// <param name="helper">The xml helper</param>
        /// <param name="path">The X Path</param>
        internal override void CreateNodes(XmlHelper helper, string path)
        {
            if (NumFmtID < 0 && !string.IsNullOrEmpty(Format))
            {
                NumFmtID = _styles._nextDfxNumFmtID++;
            }
            helper.CreateNode(path);
            SetValue(helper, path + "/@numFmtId", NumFmtID);
            SetValue(helper, path + "/@formatCode", Format);
        }
        /// <summary>
        /// If the object has any properties set
        /// </summary>
        public override bool HasValue
        {
            get 
            { 
                return !string.IsNullOrEmpty(Format) && NumFmtID!=0; 
            }
        }
        /// <summary>
        /// Clears all properties
        /// </summary>
        public override void Clear()
        {
            Format = null;
            NumFmtID = int.MinValue;
        }
        /// <summary>
        /// Clone the object
        /// </summary>
        /// <returns>A new instance of the object</returns>
        internal override DxfStyleBase Clone()
        {
            return new ExcelDxfNumberFormat(_styles, _callback) { NumFmtID = NumFmtID, Format = Format };
        }
        internal override void SetValuesFromXml(XmlHelper helper)
        {
            if (helper.ExistsNode("d:numFmt"))
            {
                NumFmtID = helper.GetXmlNodeInt("d:numFmt/@numFmtId");
                Format = helper.GetXmlNodeString("d:numFmt/@formatCode");
                if (NumFmtID < 164 && string.IsNullOrEmpty(Format))
                {
                    Format = ExcelNumberFormat.GetFromBuildInFromID(NumFmtID);
                }
            }
        }
        internal override void SetStyle()
        {
            if (_callback != null)
            {
                _callback?.Invoke(eStyleClass.Numberformat, eStyleProperty.Format, _format);
            }
        }
    }
}
