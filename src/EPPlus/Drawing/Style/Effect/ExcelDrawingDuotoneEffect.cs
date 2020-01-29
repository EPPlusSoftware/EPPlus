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
    /// A Duotune effect
    /// </summary>
    public class ExcelDrawingDuotoneEffect : XmlHelper
    {
        internal ExcelDrawingDuotoneEffect(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {

        }
        private  ExcelDrawingColorManager _color1;
        /// <summary>
        /// The first color
        /// </summary>
        public ExcelDrawingColorManager Color1
        {
            get
            {
                if (_color1 == null)
                {
                    _color1 = new ExcelDrawingColorManager(NameSpaceManager, TopNode.FirstChild, "", SchemaNodeOrder);
                }
                return _color1;
            }
        }
        private ExcelDrawingColorManager _color2;
        /// <summary>
        /// The second color
        /// </summary>
        public ExcelDrawingColorManager Color2
        {
            get
            {
                if (_color2 == null)
                {
                    _color2 = new ExcelDrawingColorManager(NameSpaceManager, TopNode.LastChild, "", SchemaNodeOrder);
                }
                return _color2;
            }
        }
    }
}
