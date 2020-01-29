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
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill
{
    /// <summary>
    /// A solid fill.
    /// </summary>
    public class ExcelDrawingSolidFill : ExcelDrawingFillBase
    {
        string[] _schemaNodeOrder;
        internal ExcelDrawingSolidFill(XmlNamespaceManager nsm, XmlNode topNode, string fillPath, string[]  schemaNodeOrder) : base(nsm, topNode, fillPath)
        {
            _schemaNodeOrder = schemaNodeOrder;
            GetXml();
        }
        /// <summary>
        /// The fill style
        /// </summary>
        public override eFillStyle Style
        {
            get
            {
                return eFillStyle.SolidFill;
            }
        }
        ExcelDrawingColorManager _color = null;

        /// <summary>
        /// The color of the fill
        /// </summary>
        public ExcelDrawingColorManager Color
        {
            get
            {
                if (_color == null)
                {
                    _color = new ExcelDrawingColorManager(_nsm, _topNode, _fillPath, _schemaNodeOrder);
                }
                return _color;
            }
        }

        internal override string NodeName
        {
            get
            {
                return "a:solidFill";
            }
        }

        internal override void SetXml(XmlNamespaceManager nsm, XmlNode node)
        {
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
            CheckTypeChange(NodeName);
            if(_color==null)
            {
                Color.SetPresetColor(ePresetColor.Black);
            }
            _color.SetXml(nsm, node);
        }
        internal override void GetXml()
        {
            
        }
        internal override void UpdateXml()
        {
        }
    }
}
