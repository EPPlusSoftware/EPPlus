using OfficeOpenXml.Utils.Extentions;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public abstract class ExcelChartExData : XmlHelper
    {
        internal ExcelChartExData(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
        }
        /// <summary>
        /// Data formula
        /// </summary>
        public string Formula 
        { 
            get
            {
                return GetXmlNodeString("cx:f");
            }
            set
            {
                SetXmlNodeString("cx:f", value);
            }
        }
        public eFormulaDirection FormulaDirection
        {
            get
            {
                return GetXmlNodeString("cx:f/@dir").ToEnum(eFormulaDirection.Column);
            }
            set
            {
                SetXmlNodeString("cx:f/@dir", value.ToEnumString());
            }
        }

        /// <summary>
        /// Dimension name formula. Return null if the element does not exist
        /// </summary>
        public string NameFormula
        {
            get
            {
                if(ExistNode("cx:nf"))
                {
                    return GetXmlNodeString("cx:nf");
                }
                else
                {
                    return null;
                }
            }
            set
            {
                SetXmlNodeString("cx:nf", value);
            }
        }
        /// <summary>
        /// Directopm for the name formula
        /// </summary>
        public eFormulaDirection? NameFormulaDirection 
        { 
            get
            {
                if (ExistNode("cx:nf"))
                {
                    return GetXmlNodeString("cx:nf/@dir").ToEnum<eFormulaDirection>(eFormulaDirection.Column);
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(value==null)
                {
                    DeleteNode("cx:nf/@dir");
                }
                else
                {
                    SetXmlNodeString("cx:nf/@dir", value.ToEnumString());
                }
            }
        }
    }
}
