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
            HeaderRowStyle = styles.GetDxf(HeaderRowDxfId);
            DataStyle = styles.GetDxf(DataDxfId);
            TotalsRowStyle = styles.GetDxf(TotalsRowDxfId);
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
        internal string HeaderRowStyleName
        {
            get
            {
                return GetXmlNodeString("@headerRowCellStyle");
            }
            set
            {
                SetXmlNodeString("@headerRowCellStyle",value);
            }
        }

        internal ExcelDxfStyle HeaderRowStyle { get; private set; }
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
        internal ExcelDxfStyle DataStyle { get; private set; }
        internal ExcelDxfStyle TotalsRowStyle { get; private set; }
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
