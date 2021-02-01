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
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill
{
    /// <summary>
    /// A gradient fill. This fill gradual transition from one color to the next.
    /// </summary>s
    public class ExcelDrawingGradientFill : ExcelDrawingFillBase
    {
        private string[] _schemaNodeOrder;
        internal ExcelDrawingGradientFill(XmlNamespaceManager nsm, XmlNode topNode, string[]  schemaNodeOrder) : base(nsm, topNode,"")
        {
            _schemaNodeOrder = schemaNodeOrder;
            GetXml();
        }

        /// <summary>
        /// The direction(s) in which to flip the gradient while tiling
        /// </summary>
        public eTileFlipMode TileFlip { get; set; }
        /// <summary>
        /// If the fill rotates along with shape.
        /// </summary>
        public bool RotateWithShape
        {
            get;
            set;
        }
        ExcelDrawingGradientFillColorList _colors = null;
        const string ColorsPath = "a:gsLst";
        /// <summary>
        /// A list of colors and their positions in percent used to generate the gradiant fill
        /// </summary>
        public ExcelDrawingGradientFillColorList Colors
        {
            get
            {
                if (_colors == null)
                {
                    _colors = new ExcelDrawingGradientFillColorList(_nsm, _topNode, ColorsPath, _schemaNodeOrder);
                }
                return _colors;
            }
        }
        /// <summary>
        /// The fill style. 
        /// </summary>
        public override eFillStyle Style
        {
            get
            {
                return eFillStyle.GradientFill;
            }
        }

        internal override string NodeName
        {
            get
            {
                return "a:gradFill";
            }
        }

        internal override void SetXml(XmlNamespaceManager nsm, XmlNode node)
        {
            if (_xml == null) InitXml(nsm, node,"");
            CheckTypeChange(NodeName);
            _xml.SetXmlNodeBool("@rotWithShape", RotateWithShape);
            if (TileFlip == eTileFlipMode.None)
            {
                _xml.DeleteNode("@flip");
            }
            else
            {
                _xml.SetXmlNodeString("@flip", TileFlip.ToString().ToLower());
            }

            if (ShadePath==eShadePath.Linear && LinearSettings.Angel!=0 && LinearSettings.Scaled==false)
            {
                _xml.SetXmlNodeAngel("a:lin/@ang", LinearSettings.Angel);
                _xml.SetXmlNodeBool("a:lin/@scaled", LinearSettings.Scaled);
            }
            else if(ShadePath != eShadePath.Linear)
            {
                _xml.SetXmlNodeString("a:path/@path", GetPathString(ShadePath));
                _xml.SetXmlNodePercentage("a:path/a:fillToRect/@b", FocusPoint.BottomOffset, true, int.MaxValue/10000);
                _xml.SetXmlNodePercentage("a:path/a:fillToRect/@t", FocusPoint.TopOffset, true, int.MaxValue / 10000);
                _xml.SetXmlNodePercentage("a:path/a:fillToRect/@l", FocusPoint.LeftOffset, true, int.MaxValue / 10000);
                _xml.SetXmlNodePercentage("a:path/a:fillToRect/@r", FocusPoint.RightOffset, true, int.MaxValue / 10000);
            }
        }

        private string GetPathString(eShadePath shadePath)
        {
            switch(shadePath)
            {
                case eShadePath.Circle:
                    return "circle";
                case eShadePath.Rectangle:
                    return "rect";
                case eShadePath.Shape:
                    return "shape";
                default:
                    throw (new ArgumentException("Unhandled ShadePath"));
            }
        }

        internal override void GetXml()
        {
            _colors = new ExcelDrawingGradientFillColorList(_xml.NameSpaceManager, _xml.TopNode, ColorsPath, _schemaNodeOrder);
            RotateWithShape = _xml.GetXmlNodeBool("@rotWithShape");
            try
            {
                var s = _xml.GetXmlNodeString("@flip");
                if (string.IsNullOrEmpty(s))
                {
                    TileFlip = eTileFlipMode.None;
                }
                else
                {
                    TileFlip = (eTileFlipMode)Enum.Parse(typeof(eTileFlipMode), s, true);
                }
            }
            catch
            {
                TileFlip = eTileFlipMode.None;
            }

            var cols = _xml.TopNode.SelectSingleNode("a:gsLst", _xml.NameSpaceManager);
            if (cols != null)
            {
                foreach (XmlNode c in cols.ChildNodes)
                {
                    var xml = XmlHelperFactory.Create(_xml.NameSpaceManager, c);
                    _colors.Add(xml.GetXmlNodeDouble("@pos") / 1000, c);
                }
            }
            var path=_xml.GetXmlNodeString("a:path/@path");
            if(!string.IsNullOrEmpty(path))
            {
                if (path == "rect") path = "rectangle";
                ShadePath = path.ToEnum(eShadePath.Linear);
            }
            else
            {
                ShadePath = eShadePath.Linear;
            }
            if(ShadePath==eShadePath.Linear)
            {
                LinearSettings = new ExcelDrawingGradientFillLinearSettings(_xml);
            }
            else
            {
                FocusPoint = new ExcelDrawingRectangle(_xml, "a:path/a:fillToRect/", 0);
            }
        }
        eShadePath _shadePath = eShadePath.Linear;
        /// <summary>
        /// Specifies the shape of the path to follow
        /// </summary>
        public eShadePath ShadePath
        {
            get
            {
                return _shadePath;
            }
            set
            {
                if(value==eShadePath.Linear)
                {
                    LinearSettings = new ExcelDrawingGradientFillLinearSettings();
                    FocusPoint = null;
                }
                else
                {
                    LinearSettings = null;
                    FocusPoint = new ExcelDrawingRectangle(50);
                }
                _shadePath = value;
            }
        }

        /// <summary>
        /// The focuspoint when ShadePath is set to a non linear value.
        /// This property is set to null if ShadePath is set to Linear
        /// </summary>
        public ExcelDrawingRectangle FocusPoint
        {
            get;
            private set;
        }
        /// <summary>
        /// Linear gradient settings.
        /// This property is set to null if ShadePath is set to Linear
        /// </summary>
        public ExcelDrawingGradientFillLinearSettings LinearSettings
        {
            get;
            private set;
        }
        internal override void UpdateXml()
        {
            if (_xml == null) CreateXmlHelper();
            SetXml(_nsm, _xml.TopNode);
            
        }
    }
}
