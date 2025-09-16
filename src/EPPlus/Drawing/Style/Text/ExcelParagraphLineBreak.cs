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
using System;
using System.Xml;
namespace OfficeOpenXml.Drawing
{
    public class ExcelParagraphLineBreak : ExcelParagraphTextRunBase
    {
        internal ExcelParagraphLineBreak(IPictureRelationDocument prd, XmlNamespaceManager ns, XmlNode topNode) : base(prd, ns, topNode)
        {
        }
        public override eParagraphRunType Type => eParagraphRunType.LineBreak;
        public override string Text
        {
            get
            {
                return Environment.NewLine;
            }
            set
            {
                throw new InvalidOperationException("Can not set the text for a line break paragraph item.");
            }
        }
    }
}
