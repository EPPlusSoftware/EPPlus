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
using System.Text.RegularExpressions;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using static OfficeOpenXml.Style.ExcelNumberFormat;

namespace OfficeOpenXml.Style
{
	/// <summary>
	/// The numberformat of the cell
	/// </summary>
	public sealed partial class ExcelNumberFormat : StyleBase
    {
        internal ExcelNumberFormat(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string Address, int index) :
            base(styles, ChangedEvent, PositionID, Address)
        {
            Index = index;
        }
        /// <summary>
        /// The numeric index fror the format
        /// </summary>
        public int NumFmtID 
        {
            get
            {
                return Index;
            }
        }
        /// <summary>
        /// The numberformat 
        /// </summary>
        public string Format
        {
            get
            {
                for(int i=0;i<_styles.NumberFormats.Count;i++)
                {
                    if(Index==_styles.NumberFormats[i].NumFmtId)
                    {
                        return _styles.NumberFormats[i].Format;
                    }
                }
                return "general";
            }
            set
            {                
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Numberformat, eStyleProperty.Format, (string.IsNullOrEmpty(value) ? "General" : value), _positionID, _address));
            }
        }

        internal override string Id
        {
            get 
            {
                return Format;
            }
        }
        /// <summary>
        /// If the numeric format is a build-in from.
        /// </summary>
        public bool BuildIn { get; private set; }

        internal static string GetFromBuildInFromID(int _numFmtId)
        {
            switch (_numFmtId)
            {
                case 0:
                    return "General";
                case 1:
                    return "0";
                case 2:
                    return "0.00";
                case 3:
                    return "#,##0";
                case 4:
                    return "#,##0.00";
                case 9:
                    return "0%";
                case 10:
                    return "0.00%";
                case 11:
                    return "0.00E+00";
                case 12:
                    return "# ?/?";
                case 13:
                    return "# ??/??";
                case 14:
                    return "mm-dd-yy";
                case 15:
                    return "d-mmm-yy";
                case 16:
                    return "d-mmm";
                case 17:
                    return "mmm-yy";
                case 18:
                    return "h:mm AM/PM";
                case 19:
                    return "h:mm:ss AM/PM";
                case 20:
                    return "h:mm";
                case 21:
                    return "h:mm:ss";
                case 22:
                    return "m/d/yy h:mm";
                case 37:
                    return "#,##0 ;(#,##0)";
                case 38:
                    return "#,##0 ;[Red](#,##0)";
                case 39:
                    return "#,##0.00;(#,##0.00)";
                case 40:
                    return "#,##0.00;[Red](#,##0.00)";
                case 45:
                    return "mm:ss";
                case 46:
                    return "[h]:mm:ss";
                case 47:
                    return "mmss.0";
                case 48:
                    return "##0.0";
                case 49:
                    return "@";
                default:
                    return string.Empty;
            }
        }
        internal static int GetFromBuildIdFromFormat(string format)
        {
            switch (format)
            {
                case "General":
                case "":
                    return 0;
                case "0":
                    return 1;
                case "0.00":
                    return 2;
                case "#,##0":
                    return 3;
                case "#,##0.00":
                    return 4;
                case "0%":
                    return 9;
                case "0.00%":
                    return 10;
                case "0.00E+00":
                    return 11;
                case "# ?/?":
                    return 12;
                case "# ??/??":
                    return 13;
                case "mm-dd-yy":
                    return 14;
                case "d-mmm-yy":
                    return 15;
                case "d-mmm":
                    return 16;
                case "mmm-yy":
                    return 17;
                case "h:mm AM/PM":
                    return 18;
                case "h:mm:ss AM/PM":
                    return 19;
                case "h:mm":
                    return 20;
                case "h:mm:ss":
                    return 21;
                case "m/d/yy h:mm":
                    return 22;
                case "#,##0 ;(#,##0)":
                    return 37;
                case "#,##0 ;[Red](#,##0)":
                    return 38;
                case "#,##0.00;(#,##0.00)":
                    return 39;
                case "#,##0.00;[Red](#,##0.00)":                      
                    return 40;
                case "mm:ss":
                    return 45;
                case "[h]:mm:ss":
                    return 46;
                case "mmss.0":
                    return 47;
                case "##0.0":
                    return 48;
                case "@":
                    return 49;
                default:
                    return int.MinValue;
            }
        }
        internal NumberFormatType _numberformatType= NumberFormatType.Unset;
        //TODO: Implement convert to DateTime if dateformat.
		internal bool IsDateFormat
        {
            get
            {
                if(_numberformatType==NumberFormatType.Unset)
                {
                    _numberformatType = GetNumberFormatType(Format);
				}
                return _numberformatType == NumberFormatType.Date;
            }
        }
        //TODO: Use to determin number of decimals in the number format.
        internal bool IsNumberFormat
		{
            get
            {
				if (_numberformatType == NumberFormatType.Unset)
				{
					_numberformatType = GetNumberFormatType(Format);
				}
				return _numberformatType == NumberFormatType.Numeric;
            }
		}
        internal static NumberFormatType GetNumberFormatType(string format)
        {
            if(string.IsNullOrEmpty(format)) return NumberFormatType.General;
            bool isInString = false;
            bool isInBracket = false;
            bool IsEscaped = false;
            bool isDate=false, isNumber=false, isText = false;
            foreach(var c in format.ToLower())
            {
                if (IsEscaped)
                {
                    IsEscaped = false;
                    continue;
				}
                else if((isInString && c!='\"') || isInBracket && c != ']')
                {
                    continue;
                }
                switch(c)
                {
                    case '\\':
                        IsEscaped = true;
                        break;
                    case '\"':
                        isInBracket=!isInString; 
                        break;
                    case '[':
                        isInBracket = true;
						break; ;
                    case ']':
						isInBracket = false;
                        break;
                    case 'y':
					case 'm':
					case 'd':
					case 'h':
					case 's':
                        isDate=true;
                        break;
                    case '#':
                    case '0':
						isNumber = true;
                        break;
                    case '@':
                        isText=true;
                        break;
				}
			}
            if(isNumber)
            {
                return NumberFormatType.Numeric;
            }
            else if (isDate)
            {
                return NumberFormatType.Date;
            }
            else if(isInString)
            {
                return NumberFormatType.String;
            }
            return NumberFormatType.General;
        }

	}
}
