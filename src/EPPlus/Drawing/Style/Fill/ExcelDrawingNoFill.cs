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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Style.Fill
{
    /// <summary>
    /// The drawing has no fill
    /// </summary>
    public class ExcelDrawingNoFill : ExcelDrawingFillBase
    {
        internal ExcelDrawingNoFill(ExcelDrawing drawing) : base ()
        {

        }
        /// <summary>
        /// The type of fill
        /// </summary>
        public override eFillStyle Style
        {
            get
            {
                return eFillStyle.NoFill;
            }
        }

        internal override string NodeName
        {
            get
            {
                return "a:noFill";
            }
        }

        internal override void GetXml()
        {

        }

        internal override void SetXml(XmlNamespaceManager nsm, XmlNode node)
        {
            if (_xml == null) InitXml(nsm, node.FirstChild, "");
            CheckTypeChange(NodeName);
        }

        internal override void UpdateXml()
        {
            
        }
    }
}
