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
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill
{
    /// <summary>
    /// A pattern fill.
    /// </summary>
    public class ExcelDrawingPatternFill : ExcelDrawingFillBase
    {
        string[] _schemaNodeOrder;
        internal ExcelDrawingPatternFill(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string fillPath, string[] schemaNodeOrder, Action initXml) : base(nameSpaceManager, topNode, fillPath, initXml)
        {
            _schemaNodeOrder = XmlHelper.CopyToSchemaNodeOrder(schemaNodeOrder, new string[] { "fgClr", "bgClr" });
            GetXml();
        }
        /// <summary>
        /// The fillstyle, always PatternFill
        /// </summary>
        public override eFillStyle Style
        {
            get
            {
                return eFillStyle.PatternFill;
            }
        }
        private eFillPatternStyle _pattern;
        /// <summary>
        /// The preset pattern to use
        /// </summary>
        public eFillPatternStyle PatternType
        {
            get
            {
                return _pattern;
            }
            set
            {
                _pattern = value;
            }
        }
        ExcelDrawingColorManager _fgColor = null;
        /// <summary>
        /// Foreground color
        /// </summary>
        public ExcelDrawingColorManager ForegroundColor
        {
            get
            {
                if (_fgColor == null)
                {
                    _fgColor = new ExcelDrawingColorManager(_nsm, _topNode, "a:fgClr", _schemaNodeOrder, _initXml);
                }
                return _fgColor;
            }
        }
        ExcelDrawingColorManager _bgColor = null;
        /// <summary>
        /// Background color
        /// </summary>
        public ExcelDrawingColorManager BackgroundColor
        {
            get
            {
                if(_bgColor == null)
                {
                    _bgColor = new ExcelDrawingColorManager(_nsm, _topNode, "a:bgClr", _schemaNodeOrder, _initXml);
                }
                return _bgColor;
            }
        }


        internal override string NodeName
        {
            get
            {
                return "a:patternFill";
            }
        }

        internal override void SetXml(XmlNamespaceManager nsm, XmlNode node)
        {
            _initXml?.Invoke();
            if (_xml == null)
            {
                if(string.IsNullOrEmpty(_fillPath))
                {
                    InitXml(nsm, node,"");
                }
                else
                {
                    CreateXmlHelper();
                }
            }
            _xml.SetXmlNodeString("@prst", PatternType.ToEnumString());
            var fgNode=_xml.CreateNode("a:fgClr");
            ForegroundColor.SetXml(nsm, fgNode);

            var bgNode = _xml.CreateNode("a:bgClr");
            BackgroundColor.SetXml(nsm, bgNode);
        }
        internal override void GetXml()
        {
            PatternType = _xml.GetXmlNodeString("@prst").ToEnum(eFillPatternStyle.Pct5);
        }

        internal override void UpdateXml()
        {
            if (_xml == null) CreateXmlHelper();
            SetXml(_nsm, _xml.TopNode);
        }
    }
}
