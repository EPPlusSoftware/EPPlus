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
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Properties for drawing line ends
    /// </summary>
    public sealed class ExcelDrawingLineEnd:XmlHelper
    {
         string _linePath;
        private readonly Action _init;
        internal ExcelDrawingLineEnd(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string linePath, Action init) : 
            base(nameSpaceManager, topNode)
        {
            _linePath = linePath;
            _init = init;
            SchemaNodeOrder = new string[] { "noFill","solidFill","gradFill","pattFill","prstDash", "custDash", "round","bevel", "miter", "headEnd", "tailEnd" };
        }
        string _stylePath = "/@type";
        /// <summary>
        /// The shapes line end decoration
        /// </summary>
        public eEndStyle? Style
        {
            get
            {
                return TranslateEndStyle(GetXmlNodeString(_linePath + _stylePath));
            }
            set
            {
                _init();
                if (value == null)
                {
                    DeleteNode(_linePath + _stylePath);
                }
                else
                {
                    SetXmlNodeString(_linePath + _stylePath, TranslateEndStyleText(value.Value));
                }
            }
        }

        string _widthPath = "/@w";
        /// <summary>
        /// The line start/end width in relation to the line width
        /// </summary>
        public eEndSize? Width
        {
            get
            {
                return TranslateEndSize(GetXmlNodeString(_linePath + _widthPath));
            }
            set
            {
                _init();
                if (value == null)
                {
                    DeleteNode(_linePath + _widthPath);
                }
                else
                {
                    SetXmlNodeString(_linePath + _widthPath, TranslateEndSizeText(value.Value));
                }
            }
        }

        string _heightPath = "/@len";
        /// <summary>
        /// The line start/end height in relation to the line height
        /// </summary>
        public eEndSize? Height
        {
            get
            {
                return TranslateEndSize(GetXmlNodeString(_linePath +_heightPath));
            }
            set
            {
                _init();
                if (value == null)
                {
                    DeleteNode(_linePath + _heightPath);
                }
                else
                {
                    SetXmlNodeString(_linePath + _heightPath, TranslateEndSizeText(value.Value));
                }
            }
        }
        #region "Translate Enum functions"
        private string TranslateEndStyleText(eEndStyle value)
        {
            return value.ToString().ToLower();
        }
        private eEndStyle? TranslateEndStyle(string text)
        {
            switch (text)
            {
                case "none":
                case "arrow":
                case "diamond":
                case "oval":
                case "stealth":
                case "triangle":
                    return (eEndStyle)Enum.Parse(typeof(eEndStyle), text, true);
                default:
                    return null;
            }
        }
        private string GetCreateLinePath(bool doCreate)
        {
            if (string.IsNullOrEmpty(_linePath))
            {
                return "";
            }
            else
            {
                if(doCreate) CreateNode(_linePath, false);
                return _linePath + "/";
            }
        }

        private string TranslateEndSizeText(eEndSize value)
        {
            string text = value.ToString();
            switch (value)
            {
                case eEndSize.Small:
                    return "sm";
                case eEndSize.Medium:
                    return "med";
                case eEndSize.Large:
                    return "lg";
                default:
                    return null;
            }
        }
        private eEndSize? TranslateEndSize(string text)
        {
            switch (text)
            {
                case "sm":
                    return eEndSize.Small;
                case "med":
                    return eEndSize.Medium;
                case "lg":
                    return eEndSize.Large;
                default:
                    return null;
            }
        }
        #endregion
    }
}