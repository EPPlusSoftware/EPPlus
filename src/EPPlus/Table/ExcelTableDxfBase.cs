using OfficeOpenXml.Style.Dxf;
using System.Xml;

namespace OfficeOpenXml.Table
{
    public class ExcelTableDxfBase : XmlHelper
    {
        internal ExcelTableDxfBase(XmlNamespaceManager nsm) : base(nsm)
        {
        }
        internal ExcelTableDxfBase(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
        }
        internal void InitStyles(ExcelStyles styles)
        {            
            HeaderRowStyle = new ExcelDxfStyle(NameSpaceManager, TopNode, styles, "@headerRowDxfId");
            DataStyle = new ExcelDxfStyle(NameSpaceManager, TopNode, styles, "@dataDxfId");
            TotalsRowStyle = new ExcelDxfStyle(NameSpaceManager, TopNode, styles, "@totalsRowDxfId");
        }
        internal int HeaderRowDxfId
        {
            get
            {
                return GetXmlNodeInt("@headerRowDxfId");
            }
            set
            {
                SetXmlNodeInt("@headerRowDxfId", value);
            }
        }
        public ExcelDxfStyle HeaderRowStyle { get; private set; }
        internal int DataDxfId
        {
            get
            {
                return GetXmlNodeInt("@dataDxfId");
            }
            set
            {
                SetXmlNodeInt("@dataDxfId", value);
            }
        }
        public ExcelDxfStyle DataStyle { get; private set; }
        public ExcelDxfStyle TotalsRowStyle { get; private set; }
        internal int TotalsRowDxfId
        {
            get
            {
                return GetXmlNodeInt("@totalsRowDxfId");
            }
            set
            {
                SetXmlNodeInt("@totalsRowDxfId", value);
            }
        }


    }
}
