/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
    9/11/2025         EPPlus Software AB       EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Describes a tab stop in a paragraphs tab stop collection.
    /// </summary>
    public class ExcelDrawingParagraphTabStop : XmlHelper
    {
        Action _initXml;
        internal ExcelDrawingParagraphTabStop(XmlNamespaceManager nsm, XmlElement topNode, string[] schemaNodeOrder, Action initXml) : base(nsm, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _initXml= initXml; 
        }
        /// <summary>
        /// How the text is aligned for a particular tab stop.
        /// </summary>
        public eTabStopParagraphAlignment Alignment
        {
            get
            {
                return GetXmlNodeString("@algn").ToEnum(eTabStopParagraphAlignment.Left, new Dictionary<string, eTabStopParagraphAlignment>
                {
                    ["r"] = eTabStopParagraphAlignment.Right,
                    ["l"] = eTabStopParagraphAlignment.Left,
                    ["dec"] = eTabStopParagraphAlignment.Decimal,
                    ["ctr"] = eTabStopParagraphAlignment.Center
                });
            }
            set
            {
                _initXml.Invoke();  
                SetXmlNodeString("@algn", value.ToEnumString(new Dictionary<Enum, string>
                {
                    [eTabStopParagraphAlignment.Right] = "r",
                    [eTabStopParagraphAlignment.Left] = "l",
                    [eTabStopParagraphAlignment.Decimal] = "dec",
                    [eTabStopParagraphAlignment.Center] = "ctr"
                }));
            }
        }
        /// <summary>
        /// The position of the tab stop relative to the left margin in pixels.
        /// </summary>
        public double Position 
        {
            get
            {
                return GetXmlNodeEmuToPixel("@pos");
            }
            set
            {
                _initXml.Invoke();
                SetXmlNodeEmuToPixel("@pos", value);
            }
        }
    }
}