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
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Border for drawings
    /// </summary>    
    public sealed class ExcelDrawingBorder : XmlHelper
    {
        string _linePath;
        IPictureRelationDocument _pictureRelationDocument;
        bool isSpInit=false;
        internal ExcelDrawingBorder(IPictureRelationDocument pictureRelationDocument, XmlNamespaceManager nameSpaceManager, XmlNode topNode, string linePath, string[] schemaNodeOrder) : 
            base(nameSpaceManager, topNode)
        {
            AddSchemaNodeOrder(schemaNodeOrder, new string[] { "noFill", "solidFill", "gradFill", "pattFill", "prstDash", "custDash", "round","bevel","miter","headEnd","tailEnd" });
            _linePath = linePath;   
            _lineStylePath = string.Format(_lineStylePath, linePath);
            _lineCapPath = string.Format(_lineCapPath, linePath);
            _lineWidth = string.Format(_lineWidth, linePath);
            _bevelPath = string.Format(_bevelPath, linePath);
            _roundPath = string.Format(_roundPath, linePath);
            _miterPath = string.Format(_miterPath, linePath);
            _miterJoinLimitPath = string.Format(_miterJoinLimitPath, linePath);
                
            _headEndPath = string.Format(_headEndPath, linePath);
            _tailEndPath = string.Format(_tailEndPath, linePath);
            _compoundLineTypePath = string.Format(_compoundLineTypePath, linePath);
            _alignmentPath = string.Format(_alignmentPath, linePath);
            _pictureRelationDocument = pictureRelationDocument;
        }

        #region "Public properties"
        ExcelDrawingFillBasic _fill = null;
        /// <summary>
        /// Access to fill properties
        /// </summary>
        public ExcelDrawingFillBasic Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFillBasic(_pictureRelationDocument.Package, NameSpaceManager, TopNode, _linePath, SchemaNodeOrder, true);
                }
                return _fill;
            }
        }
        string _lineStylePath = "{0}/a:prstDash/@val";
        /// <summary>
        /// Preset line dash
        /// </summary>
        public eLineStyle? LineStyle
        {
            get
            {
                var v = GetXmlNodeString(_lineStylePath);
                if (string.IsNullOrEmpty(v))
                {
                    return null;
                }
                else
                {
                    return EnumTransl.ToLineStyle(v);
                }
            }
            set
            {
                InitSpPr();
                CreateNode(_linePath, false);
                if(value.HasValue)
                {
                    SetXmlNodeString(_lineStylePath, EnumTransl.FromLineStyle(value.Value));
                }
                else
                {
                    DeleteNode(_lineStylePath);
                }
            }
        }

        private void InitSpPr()
        {
            if(isSpInit==false)
            {
                if(CreateNodeUntil(_linePath, "spPr", out XmlNode spPrNode))
                {
                    spPrNode.InnerXml = "<a:noFill/><a:ln><a:noFill/></a:ln ><a:effectLst/><a:sp3d/>";
                 }
            }
            isSpInit = true;
        }


        string _compoundLineTypePath = "{0}/@cmpd";
        /// <summary>
        /// The compound line type that is to be used for lines with text such as underlines
        /// </summary>
        public eCompundLineStyle CompoundLineStyle
        {
            get
            {
                return EnumTransl.ToLineCompound(GetXmlNodeString(_compoundLineTypePath));
            }
            set
            {
                InitSpPr();
                SetXmlNodeString(_compoundLineTypePath, EnumTransl.FromLineCompound(value));
            }
        }
        string _alignmentPath = "{0}/@algn";
        /// <summary>
        /// The pen alignment type for use within a text body
        /// </summary>
        public ePenAlignment Alignment
        {
            get
            {
                return EnumTransl.ToPenAlignment(GetXmlNodeString(_alignmentPath));
            }
            set
            {
                InitSpPr();
                SetXmlNodeString(_alignmentPath, EnumTransl.FromPenAlignment(value));
            }
        }
        string _lineCapPath = "{0}/@cap";
        /// <summary>
        /// Specifies how to cap the ends of lines
        /// </summary>
        public eLineCap LineCap
        {
            get
            {
                return EnumTransl.ToLineCap(GetXmlNodeString(_lineCapPath));
            }
            set
            {
                InitSpPr();
                SetXmlNodeString(_lineCapPath, EnumTransl.FromLineCap(value));
            }
        }
        string _lineWidth = "{0}/@w";
        /// <summary>
        /// Width in pixels
        /// </summary>
        public double Width
        {
            get
            {
                return GetXmlNodeEmuToPt(_lineWidth);
            }
            set
            {
                InitSpPr();
                SetXmlNodeEmuToPt(_lineWidth, value);
            }
        }
        string _bevelPath = "{0}/a:bevel";
        string _roundPath = "{0}/a:round";
        string _miterPath = "{0}/a:miter";
        /// <summary>
        /// How connected lines are joined
        /// </summary>
        public eLineJoin? Join
        {
            get
            {
                if (ExistNode(_bevelPath))
                {
                    return eLineJoin.Bevel;
                }
                else if (ExistNode(_roundPath))
                {
                    return eLineJoin.Round;
                }
                else if (ExistNode(_miterPath))
                {
                    return eLineJoin.Miter;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                InitSpPr();
                if (value == eLineJoin.Bevel)
                {
                    CreateNode(_bevelPath);
                    DeleteNode(_roundPath);
                    DeleteNode(_miterPath);
                }
                else if (value == eLineJoin.Round)
                {
                    CreateNode(_roundPath);
                    DeleteNode(_bevelPath);
                    DeleteNode(_miterPath);
                }
                else
                {
                    CreateNode(_miterPath);
                    DeleteNode(_roundPath);
                    DeleteNode(_bevelPath);
                }
            }
        }
        string _miterJoinLimitPath = "{0}/a:miter/@lim";
        /// <summary>
        /// The amount by which lines is extended to form a miter join 
        /// Otherwise miter joins can extend infinitely far.
        /// </summary>
        public double? MiterJoinLimit
        {
            get
            {
                return GetXmlNodePercentage(_miterJoinLimitPath);
            }
            set
            {
                Join = eLineJoin.Miter;
                SetXmlNodePercentage(_miterJoinLimitPath, value, false, double.MaxValue);
            }
        }
        string _headEndPath = "{0}/a:headEnd";
        ExcelDrawingLineEnd _headEnd = null;
        /// <summary>
        /// Head end style for the line
        /// </summary>
        public ExcelDrawingLineEnd HeadEnd
        {
            get
            {
                if (_headEnd == null)
                {
                    return new ExcelDrawingLineEnd(NameSpaceManager, TopNode, _headEndPath, InitSpPr);
                }
                return _headEnd;
            }
        }
        string _tailEndPath = "{0}/a:tailEnd";
        ExcelDrawingLineEnd _tailEnd = null;
        /// <summary>
        /// Tail end style for the line
        /// </summary>
        public ExcelDrawingLineEnd TailEnd
        {
            get
            {
                if (_tailEnd == null)
                {
                    return new ExcelDrawingLineEnd(NameSpaceManager, TopNode, _tailEndPath, InitSpPr);
                }
                return _tailEnd;
            }
        }

        #endregion
        internal XmlElement LineElement
        {
            get
            {
                return TopNode.SelectSingleNode(_linePath, NameSpaceManager) as XmlElement;
            }   
        }
        internal void SetFromXml(XmlElement copyFromLineElement)
        {
            InitSpPr();
            XmlElement lineElement=LineElement;
            if(lineElement==null)
            {
                CreateNode(_linePath);
            }
            foreach (XmlAttribute a in copyFromLineElement.Attributes)
            {
                lineElement.SetAttribute(a.Name, a.NamespaceURI, a.Value);
            }
            lineElement.InnerXml = copyFromLineElement.InnerXml;
        }

    }
}
