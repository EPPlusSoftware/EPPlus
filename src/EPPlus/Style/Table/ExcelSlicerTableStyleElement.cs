/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/08/2021         EPPlus Software AB       Table Styling - EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Xml;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// A style element for a custom slicer style with band
    /// </summary>
    public class ExcelSlicerTableStyleElement : XmlHelper
    {
        ExcelStyles _styles;
        internal ExcelSlicerTableStyleElement(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelStyles styles, eTableStyleElement type) : base(nameSpaceManager, topNode)
        {
            _styles = styles;
            Type = type;
        }
        ExcelDxfStyle _style = null;
        /// <summary>
        /// Access to style settings
        /// </summary>
        public ExcelDxfStyle Style
        {
            get
            {
                if (_style == null)
                {
                    _style = _styles.GetDxf(GetXmlNodeIntNull("@dxfId"));
                }
                return _style;
            }
            internal set
            {
                _style = value;
            }
        }
        /// <summary>
        /// The type of custom style element for a table style
        /// </summary>
        public eTableStyleElement Type
        {
            get;
        }
        internal virtual void CreateNode()
        {
            if(TopNode.LocalName!= "tableStyleElement")
            {
                TopNode = CreateNode("d:tableStyleElement", false, true);
            }

            SetXmlNodeString("@type", Type.ToEnumString());
            SetXmlNodeInt("@dxfId", Style.DxfId);
        }
    }
}
