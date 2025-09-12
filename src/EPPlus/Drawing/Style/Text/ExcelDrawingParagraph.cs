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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.NetworkInformation;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Represents a paragraph in a richtext within a drawing object.
    /// </summary>
    public class ExcelDrawingParagraph : XmlHelper
    {
        Action _initXml;
        IPictureRelationDocument _prd;
        internal ExcelDrawingParagraph(IPictureRelationDocument prd, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, Action initXml) : base(nameSpaceManager, topNode)
        {
            AddSchemaNodeOrder(schemaNodeOrder, ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "tabLst", "defRPr"]);
            _initXml = initXml;
            DefaultRunProperties = new ExcelTextFont(prd, nameSpaceManager, topNode, "a:pPr/a:defRPr", schemaNodeOrder, initXml);
        }
        /// <summary>
        /// Default font and fill properties for all text runs.
        /// </summary>
        public ExcelTextFont DefaultRunProperties 
        { 
            get; 
        }
        /// <summary>
        /// A collection of text runs for the paragraph
        /// </summary>
        public ExcelDrawingTextRunCollection TextRuns 
        { 
            get;  
        }
        /// <summary>
        /// Horizontal Alignment
        /// </summary>
        public eTextAlignment HorizontalAlignment
        {
            get
            {
                return GetXmlNodeString("a:pPr/@algn").ToEnum(eTextAlignment.Left, new Dictionary<string, eTextAlignment>
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
                _initXml.Invoke();
                SetXmlNodeString("a:pPr/@algn", value.ToEnumString(new Dictionary<Enum, string>
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
        /// <summary>
        /// Default width in pixels for a TAB character.
        /// </summary>
        public double? DefaultTabSize
        {
            get
            {
                return GetXmlNodeEmuToPixelNull("a:pPr/@defTabSz");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeEmuToPixel("a:pPr/@defTabSz", value);
            }
        }
        /// <summary>
        /// Left margin in pixels. This is specified in addition to the text body inset and applies only to this text paragraph
        /// </summary>
        public double? LeftMargin
        {
            get
            {
                return GetXmlNodeEmuToPixel("a:pPr/@marL", 347663 / ExcelDrawing.EMU_PER_PIXEL);
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeEmuToPixel("a:pPr/@marL", value);
            }
        }
        /// <summary>
        /// Right margin in pixels. This is specified in addition to the text body inset and applies only to this text paragraph
        /// </summary>
        public double RightMargin
        {
            get
            {
                return GetXmlNodeEmuToPixel("a:pPr/@marR");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeEmuToPixel("a:pPr/@marR", value);
            }
        }
        /// <summary>
        /// The indent size that is applied to the first line of text in the paragraph in pixels.
        /// </summary>
        public double? Indent
        {
            get
            {
                return GetXmlNodeEmuToPixelNull("a:pPr/@indent");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeEmuToPixel("a:pPr/@indent", value);
            }
        }
        /// <summary>
        /// Right-to-left flow direction.
        /// </summary>
        public bool RightToLeft
        {
            get
            {
                return GetXmlNodeBool("a:pPr/@rtl");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeBool("a:pPr/@rtl", value, false);
            }
        }
        /// <summary>
        /// The level of the paragraph in relation to the list style.
        /// </summary>
        public int Level
        {
            get
            {
                return GetXmlNodeInt("a:pPr/@lvl", 0);
            }
            set
            {
                if(value < -2 && value > 8)
                {
                    throw new ArgumentOutOfRangeException("Level must be between -2 and 8");
                }
                _initXml?.Invoke();
                SetXmlNodeInt("a:pPr/@lvl", value);
            }
        }
        /// <summary>
        /// If an Latin word can be broken in half and wrapped onto the next line without a hyphen being added.
        /// </summary>
        public bool LatinLineBreak
        {
            get
            {
                return GetXmlNodeBool("a:pPr/@latinLnBrk");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeBool("a:pPr/@latinLnBrk", value, false);
            }
        }
        /// <summary>
        /// If an East Asian word can be broken in half and wrapped onto the next line without a hyphen being added.
        /// </summary>
        public bool EastAsianLineBreak
        {
            get
            {
                return GetXmlNodeBool("a:pPr/@eaLnBrk");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeBool("a:pPr/@eaLnBrk", value, false);
            }
        }
        /// <summary>
        /// If a punctuation is to be forcefully laid out on a line of text or put on a different line of text.
        /// </summary>
        public bool HangingPunctuation 
        {
            get
            {
                return GetXmlNodeBool("a:pPr/@hangingPunct");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeBool("a:pPr/@hangingPunct", value, false);
            }
        }
        ExcelLineSpacing _spaceBefore = null;
        public ExcelLineSpacing SpaceBefore
        {
            get
            {
                if (_spaceBefore == null)
                {
                    _spaceBefore = new ExcelLineSpacing(NameSpaceManager, TopNode, "a:pPr/a:spcBef", SchemaNodeOrder, _initXml, eDrawingTextLineSpacing.Exactly);
                }
                return _spaceBefore;
            }
        }
        ExcelLineSpacing _spaceAfter = null;
        public ExcelLineSpacing SpaceAfter
        {
            get
            {
                if (_spaceAfter == null)
                {
                    _spaceAfter = new ExcelLineSpacing(NameSpaceManager, TopNode, "a:pPr/a:spcAft", SchemaNodeOrder, _initXml, eDrawingTextLineSpacing.Exactly);
                }
                return _spaceAfter;
            }
        }
        ExcelLineSpacing _lineSpacing = null;
        public ExcelLineSpacing LineSpacing
        {
            get
            {
                if (_lineSpacing == null)
                {
                    _lineSpacing = new ExcelLineSpacing(NameSpaceManager, TopNode, "a:pPr/a:lnSpc", SchemaNodeOrder, _initXml, eDrawingTextLineSpacing.Single);
                }
                return _lineSpacing;
            }
        }
        ExcelParagraphBullet _bullet = null;
        public ExcelParagraphBullet Bullet
        {
            get
            {
                if(_bullet==null)
                {
                    _bullet = new ExcelParagraphBullet(_prd, NameSpaceManager, TopNode, "a:pPr", SchemaNodeOrder, _initXml);
                }
                return _bullet;
            }
        }
        /// <summary>
        /// Vertical alignment for characters in the paragraph.
        /// </summary>
        public eTextFontAlingmentType FontAlignment
        {
            get
            {
                return GetXmlNodeString("a:pPr/@fontAlgn").ToEnum(eTextFontAlingmentType.Automatic, 
                    new Dictionary<string, eTextFontAlingmentType>
                {
                    ["t"] = eTextFontAlingmentType.Top,
                    ["b"] = eTextFontAlingmentType.Bottom,
                    ["base"] = eTextFontAlingmentType.Baseline,
                    ["ctr"] = eTextFontAlingmentType.Center,
                    ["auto"] = eTextFontAlingmentType.Automatic
                });                
            }
            set
            {
                string v = value.ToEnumString(new Dictionary<Enum, string> 
                {                    
                    [eTextFontAlingmentType.Top] = "t",
                    [eTextFontAlingmentType.Bottom] = "b",
                    [eTextFontAlingmentType.Baseline] = "base",
                    [eTextFontAlingmentType.Center] = "ctr",
                    [eTextFontAlingmentType.Automatic] = "auto",
                });

                SetXmlNodeString("a:pPr/@fontAlgn", v);
            }
        }
    }
}