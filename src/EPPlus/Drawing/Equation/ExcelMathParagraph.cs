using OfficeOpenXml;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace EPPlusTest.Drawing.Equation
{

    public enum oMathJustification
    {
        Left,
        Right,
        Center,
        CenterGroup,
    }

    public class ExcelMathParagraph : XmlHelper
    {

        internal string oMathPara = "xdr:txBody/a:p/a14:m/m:oMathPara";

        internal oMathJustification Justification
        {
            get
            {
                var jc = GetXmlNodeString("m:oMathParaPr/m:jc/@m:val");
                return jc.ToEnum<oMathJustification>(oMathJustification.CenterGroup);
            }
            set 
            {
                SetXmlNodeString("m:oMathParaPr/m:jc/@m:val", value.ToEnumString());
            }
        }

        public ExcelMathParagraph(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {

        }
    }
}
