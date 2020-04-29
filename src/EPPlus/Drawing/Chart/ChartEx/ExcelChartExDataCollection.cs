/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/29/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
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
}
