using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.Utils.Extentions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    public class ExcelChartExDataCollection : XmlHelper, IEnumerable<ExcelChartExData>
    {
        List<ExcelChartExData> _list=new List<ExcelChartExData>();
        ExcelChartExSerie _serie;
        internal ExcelChartExDataCollection(ExcelChartExSerie serie, XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
            _serie = serie;
            foreach(XmlElement c in topNode.ChildNodes)
            {
                if(c.LocalName=="numDim")
                {
                    _list.Add(new ExcelChartExNumericData(NameSpaceManager, c));
                }
                else if(c.LocalName == "strDim")
                {
                    _list.Add(new ExcelChartExStringData(NameSpaceManager, c));
                }
            }
        }
        public int Id 
        { 
            get
            {
                return GetXmlNodeInt("@id");
            }
        }
        public ExcelChartExNumericData AddNumericDimension(string formula)
        {
            var node = CreateNode("cx:numDim", true);
            var nd = new ExcelChartExNumericData(NameSpaceManager, node) { Formula = formula };
            _list.Add(nd);
            return nd;
        }
        public ExcelChartExStringData AddStringDimension(string formula)
        {
            var node = CreateNode("cx:strDim", true);
            var nd = new ExcelChartExStringData(NameSpaceManager, node) { Formula = formula };
            _list.Add(nd);
            return nd;
        }
        /// <summary>
        /// Indexer
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns></returns>
        public ExcelChartExData this[int index]
        {
            get
            {
                return _list[index];
            }
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }

        public IEnumerator<ExcelChartExData> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }

    public class ExcelChartExNumericData : ExcelChartExData
    {
        internal ExcelChartExNumericData(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
        }
        public eNumericDataType Type { get; set; }
    }
    public class ExcelChartExStringData : ExcelChartExData
    {
        internal ExcelChartExStringData(XmlNamespaceManager nsm, XmlNode topNode) : base(nsm, topNode)
        {
        }
        public eStringDataType Type { get; set; }
    }
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
