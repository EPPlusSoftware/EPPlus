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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Controls;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Handles paragraph text
    /// </summary>
    public sealed class ExcelParagraph : ExcelTextFont
    {
        internal ExcelParagraph(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager ns, XmlNode rootNode, string path, string[] schemaNodeOrder) : 
            base(pictureRelationDocument, ns, rootNode, path + "a:rPr", schemaNodeOrder)
        {
        }
        const string AligPath = "../../a:pPr/@algn";
        /// <summary>
        /// Text
        /// </summary>
        public eTextAlignment HorizontalAlignment
        {
            get
            {                
                return GetXmlNodeString(AligPath).ToEnum(eTextAlignment.Left, new Dictionary<string, eTextAlignment>
                {
                    ["r"] = eTextAlignment.Right,
                    ["ctr"] = eTextAlignment.Center,
                    ["dist"] = eTextAlignment.Distributed,
                    ["just"] = eTextAlignment.Justified,
                    ["justLow"] = eTextAlignment.JustifiedLow,
                    ["thaiDist"] = eTextAlignment.ThaiDistributed,
                    ["l"] = eTextAlignment.Left,
                }
                );
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(AligPath, value.ToEnumString(new Dictionary<Enum, string>
                {
                    [eTextAlignment.Right] = "r",
                    [eTextAlignment.Center] = "ctr",
                    [eTextAlignment.Distributed] = "dist",
                    [eTextAlignment.Justified] = "just",
                    [eTextAlignment.JustifiedLow] = "justLow",
                    [eTextAlignment.ThaiDistributed] = "thaiDist",
                    [eTextAlignment.Left] = "l",
                }));
            }
        }
        const string IndentLevelPath = "../../a:pPr/@lvl";
        /// <summary>
        /// Indent level for the paragraph. Ranges from 0-8;
        /// </summary>
        public int? IndentLevel
        {
            get
            {
                return GetXmlNodeIntNull(IndentLevelPath);
            }
            set
            {
                if(value.HasValue==false && (value < 0 || value > 8))
                {
                    throw new ArgumentOutOfRangeException("Indent level must be between 0 and 8.");
                }
                CreateTopNode();
                SetXmlNodeInt(IndentLevelPath, value);
            }
        }

        const string LineSpacingPath = "../../a:pPr/a:lnSpc";

        /// <summary>
        /// Set line spacing in Points
        /// Returns null if only defined as Percent
        /// </summary>
        public double? LineSpacingPoints
        {
            get
            {
                return GetXmlNodeIntNull(LineSpacingPath + "/a:spcPts/@val") / 100;
            }
            set
            {
                _lineSpacingType = eDrawingTextLineSpacing.Exactly;
                //the "maxInclusive value="15840000" on page 4045 of ECMA OOXML part 1
                if (value.HasValue == false && (value < 0 || value > 158400))
                {
                    throw new ArgumentOutOfRangeException("Linespacing must be between 0 and 158400 pts.");
                }
                //Poins and Percent have the same position/node and there may only be one.
                if(LineSpacingPercent != null)
                {
                    LineSpacingPercent = null;
                }
                SetXmlNodeInt(LineSpacingPath + "/a:spcPts/@val", (int)value * 100);
            }
        }
        /// <summary>
        /// Set line spacing in multiples of single lines
        /// Returns null if only defined as Exactly
        /// </summary>
        public double? LineSpacingPercent
        {
            get
            {
                return GetXmlNodePercentage(LineSpacingPath + "/a:spcPct/@val");
            }
            set
            {
                if (value.HasValue == false && (value < 0 || value > 13200))
                {
                    throw new ArgumentOutOfRangeException("Linespacing in percent must be between 0 and 13200%");
                }
                //Poins and Percent have the same position/node and there may only be one.
                if (LineSpacingPoints != null)
                {
                    LineSpacingPoints = null;
                }

                if(value == 1)
                {
                    _lineSpacingType = eDrawingTextLineSpacing.Single;
                }
                else if(value == 1.5)
                {
                    _lineSpacingType = eDrawingTextLineSpacing.OneAndAHalf;
                }
                else if(value == 2)
                {
                    _lineSpacingType = eDrawingTextLineSpacing.Double;
                }

                SetXmlNodePercentage(LineSpacingPath + "/a:spcPct/@val", value);
            }
        }

       private eDrawingTextLineSpacing _lineSpacingType;

        /// <summary>
        /// If setting Exactly or Multiple it is recommended to use
        /// LineSpacingExactly or LineSpacingMultiple propeties.
        /// Otherwise they are set to default values 13.2 or 3
        /// 
        /// Note that Single, OneAndAHalf and Double are all techically just Multiple for values 1, 1,5 and 2
        /// </summary>
        public eDrawingTextLineSpacing LineSpacing
        {
            get
            {
                return _lineSpacingType;
            }
            set
            {
                switch (value)
                {
                    case eDrawingTextLineSpacing.Single:
                        LineSpacingPercent = 1;
                        break;
                    case eDrawingTextLineSpacing.OneAndAHalf:
                        LineSpacingPercent = 1.5;
                        break;
                    case eDrawingTextLineSpacing.Double:
                        LineSpacingPercent = 2;
                        break;
                    case eDrawingTextLineSpacing.Exactly:
                        LineSpacingPoints = 13;
                        break;
                    case eDrawingTextLineSpacing.Multiple:
                        LineSpacingPercent = 3;
                        break;
                }
                _lineSpacingType = value;
            }
        }


        const string TextPath = "../a:t";
        /// <summary>
        /// Text
        /// </summary>
        public string Text
        {
            get
            {
                return GetXmlNodeString(TextPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(TextPath, value);
            }
        }
        
        /// <summary>
        /// If the paragraph is the first in the collection
        /// </summary>
        public bool IsFirstInParagraph
        {
            get
            {
                var parent = _rootNode.ParentNode;
                for (int i=0;i<parent.ChildNodes.Count;i++)
                {
                    if (parent.ChildNodes[i].LocalName == "r")
                    {
                        return parent.ChildNodes[i] == _rootNode;
                    }
                }
                return false;
            }
        }
        /// <summary>
        /// If the paragraph is the last in the collection
        /// </summary>
        public bool IsLastInParagraph
        {
            get
            {
                var parent = _rootNode.ParentNode;
                for (int i = parent.ChildNodes.Count-1; i >=0 ; i--)
                {
                    if (parent.ChildNodes[i].LocalName == "r")
                    {
                        return parent.ChildNodes[i] == _rootNode;
                    }
                }
                return false;
            }
        }
    }
}
