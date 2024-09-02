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
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml
{
    internal class ExcelPhoneticProperties : XmlHelper
    {
        internal ExcelPhoneticProperties(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
                
        }
        const string FontIdPath = "d:phoneticPr/@fontId";
        public int FontId 
        {
            get
            {
                return GetXmlNodeInt(FontIdPath);
            }
            set
            {
                SetXmlNodeInt(FontIdPath, value);
            }
        }
        const string PhoneticTypePath = "d:phoneticPr/@type";
        public ePhoneticType PhoneticType
        {
            get
            {
                return GetXmlNodeString(PhoneticTypePath).ToEnum(ePhoneticType.FullWidthKatakana);
            }
            set
            {
                SetXmlNodeString(PhoneticTypePath, GetPhoneticTypeString(value));
            }
        }
        const string PhoneticAlignmentPath = "d:phoneticPr/@alignment"; 
        public ePhoneticAlignment Alignment
        {
            get
            {
                return GetXmlNodeString(PhoneticAlignmentPath).ToEnum(ePhoneticAlignment.Left);
            }
            set
            {
                SetXmlNodeString(PhoneticTypePath, value.ToEnumString());
            }
        }
        private string GetPhoneticTypeString(ePhoneticType value)
        {
            switch (value) 
            {
                case ePhoneticType.FullWidthKatakana:
                    return "fullwidthKatakana";
                case ePhoneticType.HalfWidthKatakana:
                    return "halfwidthKatakana";
                case ePhoneticType.Hiragana:
                    return "Hiragana";
                default:
                    return "noConversion";
            }
        }
    }
    internal enum ePhoneticType
    {
        HalfWidthKatakana,
        FullWidthKatakana,
        Hiragana,
        NoConversion
    }
    internal enum ePhoneticAlignment
    {
        NoControl,
        Left,
        Center,
        Distributed
    }
}