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
        /// <list type="table">
        /// <listheader>Build in ID's</listheader>
        /// <item>0   General</item>
        /// <item>1   0</item> 
        /// <item>2   0.00</item> 
        /// <item>3   #,##0</item> 
        /// <item>4   #,##0.00</item> 
        /// <item>9   0%</item> 
        /// <item>10  0.00%</item> 
        /// <item>11  0.00E+00</item> 
        /// <item>12  # ?/?</item> 
        /// <item>13  # ??/??</item> 
        /// <item>14  mm-dd-yy</item> 
        /// <item>15  d-mmm-yy</item> 
        /// <item>16  d-mmm</item> 
        /// <item>17  mmm-yy</item> 
        /// <item>18  h:mm AM/PM</item> 
        /// <item>19  h:mm:ss AM/PM</item> 
        /// <item>20  h:mm</item> 
        /// <item>21  h:mm:ss</item> 
        /// <item>22  m/d/yy h:mm</item> 
        /// <item>37  #,##0 ;(#,##0)</item> 
        /// <item>38  #,##0 ;\[Red\](#,##0)</item> 
        /// <item>39  #,##0.00;(#,##0.00)</item> 
        /// <item>40  #,##0.00;\[Red\](#,##0.00)</item> 
        /// <item>45  mm:ss</item> 
        /// <item>46  \[h\]:mm:ss</item> 
        /// <item>47  mmss.0</item> 
        /// <item>48  ##0.0E+0</item> 
        /// <item>49  </item>@
        /// </list>
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
                NumFmtID = ExcelNumberFormat.GetIdByEnglishFormat(value);
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
		internal static string GetEmptyId()
		{
			return $"";
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
                    Format = ExcelNumberFormat.GetFormatById(NumFmtID);
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
