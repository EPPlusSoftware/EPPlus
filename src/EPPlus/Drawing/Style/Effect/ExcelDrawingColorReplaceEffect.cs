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
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Effect
{
    /// <summary>
    /// A color change effect
    /// </summary>
    public class ExcelDrawingColorReplaceEffect : XmlHelper
    {
        internal ExcelDrawingColorReplaceEffect(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {

        }
        private  ExcelDrawingColorManager _color;
        /// <summary>
        /// The color to replace with
        /// </summary>
        public ExcelDrawingColorManager Color
        {
            get
            {
                if (_color == null)
                {
                    _color = new ExcelDrawingColorManager(NameSpaceManager, TopNode, "", SchemaNodeOrder);
                }
                return _color;
            }
        }
    }
}
