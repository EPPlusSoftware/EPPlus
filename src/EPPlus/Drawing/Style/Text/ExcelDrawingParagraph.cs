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
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils;
using System;
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
        internal ExcelDrawingParagraph(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string[] schemaNodeOrder, Action initXml) : base(nameSpaceManager, topNode)
        {
            AddSchemaNodeOrder(schemaNodeOrder, ["lnSpc", "spcBef", "spcAft", "buClrTx", "buClr", "tabLst", "defRPr"]);
            _initXml = initXml;
            DefaultRunProperties = new ExcelTextFont(pictureRelationDocument, nameSpaceManager, topNode, "a:pPr/a:defRPr", schemaNodeOrder, initXml);
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
        const string AligPath = "a:pPr/@align";
        /// <summary>
        /// Horizontal Alignment
        /// </summary>
        public eTextAlignment HorizontalAlignment
        {
            get
            {
                return GetXmlNodeString(AligPath).ToEnum<eTextAlignment>(eTextAlignment.Left);
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeString(AligPath, value.ToEnumString());
            }
        }    
        /// <summary>
        /// Default width in pixels for a TAB character.
        /// </summary>
        public double? DefaultTabSize
        {
            get
            {
                return GetXmlNodeEmuToPixelNull("@defTabSz");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeEmuToPixel("@defTabSz", value);
            }
        }
        /// <summary>
        /// Left margin in pixels. This is specified in addition to the text body inset and applies only to this text paragraph
        /// </summary>
        public double? LeftMargin
        {
            get
            {
                return GetXmlNodeEmuToPixel("@marL", 347663 / ExcelDrawing.EMU_PER_PIXEL);
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeEmuToPixel("@marL", value);
            }
        }
        /// <summary>
        /// Right margin in pixels. This is specified in addition to the text body inset and applies only to this text paragraph
        /// </summary>
        public double RightMargin
        {
            get
            {
                return GetXmlNodeEmuToPixel("@marR");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeEmuToPixel("@marR", value);
            }
        }
        /// <summary>
        /// The indent size that is applied to the first line of text in the paragraph in pixels.
        /// </summary>
        public double? Indent
        {
            get
            {
                return GetXmlNodeEmuToPixelNull("@indent");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeEmuToPixel("@indent", value);
            }
        }
        /// <summary>
        /// Right-to-left flow direction.
        /// </summary>
        public bool RightToLeft
        {
            get
            {
                return GetXmlNodeBool("@rtl");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeBool("@rtl", value, false);
            }
        }
        /// <summary>
        /// The level of the paragraph in relation to the list style.
        /// </summary>
        public int Level
        {
            get
            {
                return GetXmlNodeInt("@lvl");
            }
            set
            {
                if(value < -2 && value > 8)
                {
                    throw new ArgumentOutOfRangeException("Level must be between -2 and 8");
                }
                _initXml?.Invoke();
                SetXmlNodeInt("@lvl", value);
            }
        }
        /// <summary>
        /// If an Latin word can be broken in half and wrapped onto the next line without a hyphen being added.
        /// </summary>
        public bool LatinLineBreak
        {
            get
            {
                return GetXmlNodeBool("@latinLnBrk");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeBool("@latinLnBrk", value, false);
            }
        }
        /// <summary>
        /// If an East Asian word can be broken in half and wrapped onto the next line without a hyphen being added.
        /// </summary>
        public bool EastAsianLineBreak
        {
            get
            {
                return GetXmlNodeBool("@eaLnBrk");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeBool("@eaLnBrk", value, false);
            }
        }
        /// <summary>
        /// If a punctuation is to be forcefully laid out on a line of text or put on a different line of text.
        /// </summary>
        public bool HangingPunctuation 
        {
            get
            {
                return GetXmlNodeBool("@hangingPunct");
            }
            set
            {
                _initXml?.Invoke();
                SetXmlNodeBool("@hangingPunct", value, false);
            }
        }
        ExcelLineSpacing _spaceBefore = null;
        public ExcelLineSpacing SpaceBefore
        {
            get
            {
                if (_spaceBefore == null)
                {
                    _spaceBefore = new ExcelLineSpacing(NameSpaceManager, TopNode, "a:spcBef", SchemaNodeOrder, _initXml, eDrawingTextLineSpacing.Exactly);
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
                    _spaceAfter = new ExcelLineSpacing(NameSpaceManager, TopNode, "a:spcAft", SchemaNodeOrder, _initXml, eDrawingTextLineSpacing.Exactly);
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
                    _lineSpacing = new ExcelLineSpacing(NameSpaceManager, TopNode, "a:lnSpc", SchemaNodeOrder, _initXml, eDrawingTextLineSpacing.Single);
                }
                return _lineSpacing;
            }
        }
        /// <summary>
        /// Vertical alignment for characters in the paragraph.
        /// </summary>
        public eTextFontAlingmentType FontAlignment 
        { 
            get
            {
                switch(GetXmlNodeString("@fontAlgn"))
                {
                    case "t":
                        return eTextFontAlingmentType.Top;
                    case "b":
                        return eTextFontAlingmentType.Bottom;
                    case "base":
                        return eTextFontAlingmentType.Baseline;
                    case "ctr":
                        return eTextFontAlingmentType.Center;
                    default:
                        return eTextFontAlingmentType.Automatic;
                }
            }
            set
            {
                string v;
                switch (value)
                {
                    case eTextFontAlingmentType.Top:
                        v = "t";
                        break;
                    case eTextFontAlingmentType.Bottom:
                        v = "b";
                        break;
                    case eTextFontAlingmentType.Baseline:
                        v = "base";
                        break;
                    case eTextFontAlingmentType.Center:
                        v = "ctr";
                        break;
                    default:
                        v = "auto";
                        break;
                }
                SetXmlNodeString("@fontAlgn", v);
            }
        }
    }
}