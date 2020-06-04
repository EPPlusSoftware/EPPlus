using OfficeOpenXml.Utils.Extentions;
using System;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// The data used as source for the chart. Only spreadsheet internal data is supported at this point.
    /// </summary>
    public abstract class ExcelChartExData : XmlHelper
    {
        string _worksheetName;
        internal ExcelChartExData(string worksheetName, XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
            _worksheetName = worksheetName;
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
                if (ExcelCellBase.IsValidAddress(value))
                {
                    SetXmlNodeString("cx:f", ExcelCellBase.GetFullAddress(_worksheetName, value));
                }
                else
                {
                    SetXmlNodeString("cx:f", value);
                }
            }
        }
        /// <summary>
        /// The direction of the formula
        /// </summary>
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
        /// The dimensions name formula. Return null if the element does not exist
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
                if(ExcelCellBase.IsValidAddress(value))
                {
                    SetXmlNodeString("cx:nf", ExcelCellBase.GetFullAddress(_worksheetName, value), true);
                }
                else
                {
                    SetXmlNodeString("cx:nf", value, true);
                }
            }
        }
        /// <summary>
        /// Direction for the name formula
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
