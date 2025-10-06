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
using OfficeOpenXml.Core.Worksheet.Fonts.TrueTypeFontMetrics;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Text;
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
using static System.Net.Mime.MediaTypeNames;

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
        internal ExcelDrawingParagraph(ExcelDrawingParagraphCollection paragraphs, IPictureRelationDocument prd, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, Action initXml) : base(nameSpaceManager, topNode)
        {
            _paragraphs = paragraphs;
            AddSchemaNodeOrder(schemaNodeOrder, ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "buSzPct", "buSzTx", "buSzPts", "buFont", "buFontTx", "buAutoNum", "buChar", "buBlip", "buNone", "tabLst", "defRPr"]);
            _initXml = initXml;
            _prd = prd;

            //if(_paragraphs.FirstDefaultRunProperties != null)
            //{
            //    DefaultRunProperties = _paragraphs.FirstDefaultRunProperties;
            //}
            //else
            //{

            //}

            DefaultRunProperties = new ExcelTextFontXml(prd, nameSpaceManager, topNode, "a:pPr/a:defRPr", schemaNodeOrder, initXml);
            var normalStyle = _prd.Package.Workbook.Styles.GetNormalStyle();

            //////Previously new paragraphs used the first DefaultRunProperties
            //////Uncertain if we should keep this behaviour at least as an option. TODO: Decide if breaking change or legacy setting (or keep only previous paragraph's settings?)
            bool legacyDefaultRunPropertySetting = true;

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

            //var font = DefaultRunProperties.LatinFont;
            //    parentNode = CreateNode(_path);
            //    _paragraphs.Add((XmlElement)parentNode);
            //    var defNode = CreateNode(_path + "/a:pPr/a:defRPr");
            //    if (defNode.InnerXml == "")
            //    {
            //        ((XmlElement)defNode).SetAttribute("sz", (_defaultFontSize*100).ToString(CultureInfo.InvariantCulture));
            //        var normalStyle = _drawing._drawings.Worksheet.Workbook.Styles.GetNormalStyle();
            //        if (normalStyle == null)
            //            defNode.InnerXml = "<a:latin typeface=\"Calibri\" /><a:cs typeface=\"Calibri\" />";
            //        else
            //            defNode.InnerXml = $"<a:latin typeface=\"{normalStyle.Style.Font.Name}\"/><a:cs typeface=\"{normalStyle.Style.Font.Name}\"/>";
            //    }

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
        /// 0 means indent is same as MarL attribute.If this node would be null then it is considered -342900 (to counter the default value of MarL)
        /// </summary>
        public double Indent
        {
            get
            {
                return GetXmlNodeEmuToPixel("a:pPr/@marL", -342900 / ExcelDrawing.EMU_PER_PIXEL);
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
                if (value < -2 && value > 8)
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
        /// <summary>
        /// Get paragraph lineSpacing in points
        /// </summary>
        /// <param name="measurer"></param>
        /// <param name="isFirstLine">The first line in a paragraph collection has special lineHeight</param>
        /// <returns></returns>
        private double GetParagraphLineSpacing(ITextMeasurerWrap measurer, bool isFirstLine)
        {
            if (LineSpacing.LineSpacingType == eDrawingTextLineSpacing.Exactly)
            {
                return LineSpacing.Value;
            }
            else
            {
                var multiplier = LineSpacing.Value / 100;
                if (isFirstLine && this == _paragraphs[0])
                {
                    return multiplier * measurer.GetBaseLine();
                }

                return multiplier * measurer.GetSingleLineSpacing();
            }
        }

        /// <summary>
        /// Returns paragraph height in points
        /// </summary>
        /// <param name="measurer">The wrapping textMeasurer to use</param>
        /// <param name="maxWidth">MaxWidth/Wrapping width in points. A value of 0 implies no wrapping.</param>
        /// <returns></returns>
        internal double GetParagraphHeight(ITextMeasurerWrap measurer, double maxWidth = 0)
        {
            double paragraphHeight = 0;

            bool isFirstLine = true;

            foreach (var txtRun in TextRuns)
            {
                //Split textrun text into line-breaks
                var lines = txtRun.SplitIntoLines();
                //For each line in each linebreak
                foreach (var line in lines)
                {
                    var measurementFont = txtRun.GetMeasurementFont();
                    //Get the length/height of the line via the font of the textRun
                    var measurement = measurer.MeasureText(line, measurementFont);

                    //If text wrapping is on each of the broken lines could potentially be wrapped
                    List<string> finalLines = new List<string>();
                    if (maxWidth != 0)
                    {
                        var maxWidthInPixels = (maxWidth / 72d) * 96d;
                        finalLines = measurer.MeasureAndWrapText(line, measurementFont, maxWidthInPixels);
                    }
                    else
                    {
                        finalLines.Add(line);
                    }


                    //Could be just one line or mutliple lines.
                    //Re-use same collection to avoid code repetition.
                    //Line-spacing should be applied for each line
                    foreach (var fLine in finalLines)
                    {
                        //MeasureText sets the font allowing for getting the font-specific line-spacing for the text-run if it is of multiple type.
                        var lineSpacing = GetParagraphLineSpacing(measurer, isFirstLine);
                        paragraphHeight += lineSpacing;
                        isFirstLine = false;
                    }
                }
            }

            return paragraphHeight;
        }

        /// <summary>
        /// </summary>
        /// <param name="measurer"></param>
        /// <param name="maxWidth">must be entered in points</param>
        /// <returns></returns>
        internal double GetParagraphHeightInPixels(ITextMeasurerWrap measurer, double maxWidth = 0)
        {
            var pointHeight = GetParagraphHeight(measurer, maxWidth);
            var pixelHeight = (pointHeight / 72d) * 96d;

            return pixelHeight;
        }

        /// <summary>
        /// Returns paragraph height in points
        /// </summary>
        /// <param name="maxWidth">MaxWidth/Wrapping width in points. A value of 0 implies no wrapping.</param>
        /// <returns></returns>
        internal RectBase GetParagraphSize(double maxWidth = 0, double maxHeight = 0)
        {
            double paragraphHeight = 0;
            double paragraphWidth = 0;
            var measurer = _prd.Package.Settings.TextSettings.GenericTextMeasurerTrueType;
            bool isFirstLine = true;

            var maxWidthInPixels = (maxWidth / 72d) * 96d;

            foreach (var txtRun in TextRuns)
            {
                //Split textrun text into line-breaks
                var lines = txtRun.SplitIntoLines();
                //For each line in each linebreak
                foreach (var line in lines)
                {
                    var measurementFont = txtRun.GetMeasureFont();
                    //Get the length/height of the line via the font of the textRun

                    var measurement = measurer.MeasureText(line, measurementFont);

                    //If text wrapping is on each of the broken lines could potentially be wrapped
                    List<string> finalLines = new List<string>();
                    if (maxWidth != 0)
                    {
                        finalLines = measurer.MeasureAndWrapText(line, txtRun.GetMeasureFont(), maxWidthInPixels);
                    }
                    else
                    {
                        finalLines.Add(line);
                    }
                    if (measurement.Width > paragraphWidth) paragraphWidth = measurement.Width;
                    //Could be just one line or mutliple lines.
                    //Re-use same collection to avoid code repetition.
                    //Line-spacing should be applied for each line
                    foreach (var fLine in finalLines)
                    {
                        //MeasureText sets the font allowing for getting the font-specific line-spacing for the text-run if it is of multiple type.
                        var lineSpacing = GetParagraphLineSpacing(measurer, isFirstLine);
                        paragraphHeight += lineSpacing;
                        isFirstLine = false;
                    }
                }
            }

            return new RectBase(paragraphWidth, paragraphHeight);
        }

        /// <summary>
        /// </summary>
        /// <param name="maxWidth">must be entered in points</param>
        /// <returns></returns>
        internal RectBase GetParagraphSizeInPixels(double maxWidth = 0, double maxHeight = 0)
        {
            var pointHeight = GetParagraphSize(maxWidth, maxHeight);
            return new RectBase((pointHeight.Width / 72d) * 96d, (pointHeight.Height / 72d) * 96d);
        }
    }
}
