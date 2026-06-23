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
using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils; 
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Represents a paragraph in a richtext within a drawing object.
    /// </summary>
    public class ExcelDrawingParagraph : XmlHelper
    {
        Action _initXml;
        internal IPictureRelationDocument _prd;
        internal ExcelDrawingParagraphCollection _paragraphs;

        //bool legacyDefaultRunPropertySetting = false;

        internal ExcelDrawingParagraph(ExcelDrawingParagraphCollection paragraphs, IPictureRelationDocument prd, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, Action initXml) : base(nameSpaceManager, topNode)
        {
            _paragraphs = paragraphs;
            AddSchemaNodeOrder(schemaNodeOrder, ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzPct", "buSzTx", "buSzPts", "buFont", "buFontTx", "buAutoNum", "buChar", "buBlip", "buNone", "tabLst", "defRPr"]);
            _initXml = initXml;
            _prd = prd;



            if (_paragraphs.FirstDefaultRunProperties == null)
            {
                DefaultRunProperties = new ExcelTextFontXml(prd, nameSpaceManager, topNode, "a:pPr/a:defRPr", schemaNodeOrder, initXml);
            }
            else
            {
                if(paragraphs.Count == 0)
                {
                    //The node must still be created
                    var xmlFirstDefault = ((ExcelTextFontXml)paragraphs.FirstDefaultRunProperties).XmlHelper;
                    var textFont = new ExcelTextFontXml(prd, nameSpaceManager, topNode, "a:pPr/a:defRPr", schemaNodeOrder, initXml);
                    var xmlNewNode = textFont.XmlHelper;
                    CopyElement((XmlElement)xmlFirstDefault.TopNode, (XmlElement)xmlNewNode.TopNode);
                    DefaultRunProperties = textFont;
                }
                else
                {
                    DefaultRunProperties = _paragraphs.FirstDefaultRunProperties;
                }
            }

            var normalStyle = _prd.Package.Workbook.Styles.GetNormalStyle();

            //////Previously new paragraphs used the first DefaultRunProperties
            //////Uncertain if we should keep this behaviour at least as an option. TODO: Decide if breaking change or legacy setting (or keep only previous paragraph's settings?)
            bool legacyDefaultRunPropertySetting = false;

            if (paragraphs.Count == 0)
            {
                if (_paragraphs.FirstDefaultRunProperties == null)
                {
                    if (normalStyle == null)
                    {
                        DefaultRunProperties.LatinFont = DefaultRunProperties.ComplexFont = "Calibri";
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(DefaultRunProperties.LatinFont))
                        {
                            DefaultRunProperties.LatinFont = normalStyle.Style.Font.Name;
                        }

                        if (string.IsNullOrEmpty(DefaultRunProperties.ComplexFont))
                        {
                            DefaultRunProperties.ComplexFont = normalStyle.Style.Font.Name;
                        }
                    }
                }
            }
            else if (legacyDefaultRunPropertySetting && _paragraphs.FirstDefaultRunProperties != null)
            {
                ((ExcelTextFontXml)DefaultRunProperties).TriggerCreateTopNodeOnTextSet();

                var xmlFirstDefault = ((ExcelTextFontXml)paragraphs.FirstDefaultRunProperties).XmlHelper;
                var xmlNewNode = ((ExcelTextFontXml)DefaultRunProperties).XmlHelper;
                CopyElement((XmlElement)xmlFirstDefault.TopNode, (XmlElement)xmlNewNode.TopNode);
            }
        }
        /// <summary>
        /// Default font and fill properties for all text runs.
        /// </summary>
        public ExcelTextFont DefaultRunProperties
        {
            get;
        }

        ExcelDrawingTextRunCollection _textRun = null;
        /// <summary>
        /// A collection of text runs for the paragraph
        /// </summary>
        public ExcelDrawingTextRunCollection TextRuns
        {
            get
            {
                if (_textRun == null)
                {
                    _textRun = new ExcelDrawingTextRunCollection(this, NameSpaceManager, TopNode, _initXml);
                }
                return _textRun;
            }
        }
        /// <summary>
        /// The text for the paragraph.
        /// </summary>
        public string Text
        {
            get
            {
                var sb = new StringBuilder();
                foreach (var tr in TextRuns)
                {
                    sb.Append(tr.Text);
                }
                return sb.ToString();
            }
        }


        internal eTextAlignment defaultAlignment = eTextAlignment.Left;
        /// <summary>
        /// Horizontal Alignment
        /// </summary>
        public eTextAlignment HorizontalAlignment
        {
            get
            {
                return GetXmlNodeString("a:pPr/@algn").ToEnum(defaultAlignment, new Dictionary<string, eTextAlignment>
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
                _initXml?.Invoke();
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
        //If omitted default in office is 914400 EMUs
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
        public double LeftMargin
        {
            get
            {
                return GetXmlNodeEmuToPixel("a:pPr/@marL"/*,347663 / ExcelDrawing.EMU_PER_PIXEL*/);
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
        /// 0 means indent is same as MarL attribute. If this node would be null then it is considered -342900 (to counter the default value of MarL)
        /// </summary>
        public double Indent
        {
            get
            {
                var indent = GetXmlNodeEmuToPixel("a:pPr/@indent");
                //if (indent == 0)
                //{ 
                //    return GetXmlNodeEmuToPixel("a:pPr/@marL", -342900 / ExcelDrawing.EMU_PER_PIXEL);
                //}
                return indent;
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
        public int IndentLevel
        {
            get
            {
                return GetXmlNodeInt("a:pPr/@lvl", 0);
            }
            set
            {
                //From the docs:
                //"-1 and -2 for outline mode levels that should only exist in memory"
                //ECMA december_2016 part1 page 20.1.10.71 ST_TextIndentLevelType (Text Indent Level Type)
                if (value < -2 && value > 8)
                {
                    throw new ArgumentOutOfRangeException("Level must be between 0 and 8");
                }
                _initXml?.Invoke();
                SetXmlNodeInt("a:pPr/@lvl", value);
            }
        }
        /// <summary>
        /// If a Latin word can be broken in half and wrapped onto the next line without a hyphen being added.
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
                if (_bullet == null)
                {
                    _bullet = new ExcelParagraphBullet(_prd, NameSpaceManager, TopNode, "a:pPr", SchemaNodeOrder, _initXml);
                }
                return _bullet;
            }
        }
        ExcelDrawingParagraphTabStopCollection _tabStops = null;
        public ExcelDrawingParagraphTabStopCollection TabStops
        {
            get
            {
                if (_tabStops == null)
                {
                    _tabStops = new ExcelDrawingParagraphTabStopCollection(NameSpaceManager, TopNode, SchemaNodeOrder, _initXml);
                }
                return _tabStops;
            }
        }
    }
}
