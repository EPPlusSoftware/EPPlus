/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/15/2025         EPPlus Software AB       EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Drawing.Style.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils.EnumUtils;
using System;
using System.Drawing;
using System.Reflection.Emit;
using System.Xml;
namespace OfficeOpenXml.Drawing
{

    /// <summary>
    /// A regular text run.
    /// </summary>
    public class ExcelParagraphTextRun : ExcelParagraphTextRunBase
    {
        internal ExcelParagraphTextRun(ExcelDrawingParagraph paragraph, XmlNamespaceManager ns, XmlNode topNode) : base(paragraph, ns, topNode)
        {
        }
        public override eParagraphRunType Type => eParagraphRunType.TextRun;
        public override string Text
        {
            get
            {
                return GetXmlNodeString("a:t");
            }
            set
            {
                //The node must always exist therefore false.
                SetXmlNodeString("a:t", value, false);
            }
        }
    }
}
