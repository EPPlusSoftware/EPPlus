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
    public class ExcelDrawingColorChangeEffect : XmlHelper
    {
        internal ExcelDrawingColorChangeEffect(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {

        }
        private  ExcelDrawingColorManager _colorFrom;
        /// <summary>
        /// The color to transform from
        /// </summary>
        public ExcelDrawingColorManager ColorFrom
        {
            get
            {
                if (_colorFrom == null)
                {
                    var node = CreateNode("a:clrFrom");
                    _colorFrom = new ExcelDrawingColorManager(NameSpaceManager, node, "", SchemaNodeOrder);
                }
                return _colorFrom;
            }
        }
        private ExcelDrawingColorManager _colorTo;
        /// <summary>
        /// The color to transform to
        /// </summary>
        public ExcelDrawingColorManager ColorTo
        {
            get
            {
                if (_colorTo == null)
                {
                    var node = CreateNode("a:clrTo");
                    _colorTo = new ExcelDrawingColorManager(NameSpaceManager, node, "", SchemaNodeOrder);
                }
                return _colorTo;
            }

        }
    }
}
