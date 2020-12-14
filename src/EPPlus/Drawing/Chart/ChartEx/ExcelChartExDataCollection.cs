/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/29/2020         EPPlus Software AB       EPPlus 5.2
 *************************************************************************************************/
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.Utils.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A collection of chart data.
    /// </summary>
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
                    _list.Add(new ExcelChartExNumericData(serie._chart.WorkSheet.Name, NameSpaceManager, c));
                }
                else if(c.LocalName == "strDim")
                {
                    _list.Add(new ExcelChartExStringData(serie._chart.WorkSheet.Name, NameSpaceManager, c));
                }
            }
        }
        /// <summary>
        /// The id of the data
        /// </summary>
        public int Id 
        { 
            get
            {
                return GetXmlNodeInt("@id");
            }
        }
        /// <summary>
        /// Adds a numeric dimension
        /// </summary>
        /// <param name="formula">The formula or address</param>
        /// <returns>The numeric data</returns>
        public ExcelChartExNumericData AddNumericDimension(string formula)
        {
            var node = CreateNode("cx:numDim", false, true);
            var nd = new ExcelChartExNumericData(_serie._chart.WorkSheet.Name, NameSpaceManager, node) { Formula = formula };
            _list.Add(nd);
            return nd;
        }
        /// <summary>
        /// Adds a string dimension
        /// </summary>
        /// <param name="formula">The formula or address</param>
        /// <returns>The string data</returns>
        public ExcelChartExStringData AddStringDimension(string formula)
        {
            var node = CreateNode("cx:strDim", false, true);
            var nd = new ExcelChartExStringData(_serie._chart.WorkSheet.Name, NameSpaceManager, node) { Formula = formula };
            _list.Add(nd);
            return nd;
        }
        internal void SetTypeNumeric(int index, eNumericDataType type)
        {
            if(index < 0 || index >= _list.Count) throw (new IndexOutOfRangeException("index is out of range"));
            if (_list[index] is ExcelChartExStringData data)
            {
                var node = data.TopNode;
                var innerXml = data.TopNode.InnerXml;
                node.ParentNode.RemoveChild(node);

                var newNode = CreateNode("cx:numDim", false, true);
                newNode.InnerXml = innerXml;
                var nd = new ExcelChartExNumericData(_serie._chart.WorkSheet.Name, NameSpaceManager, newNode);
                nd.Type = type;
                _list[index] = nd;
            }
            else
            {
                ((ExcelChartExNumericData)_list[index]).Type = type;
            }
        }
        internal void SetTypeString(int index, eStringDataType type)
        {
            if (index < 0 || index >= _list.Count) throw (new IndexOutOfRangeException("index is out of range"));
            if (_list[index] is ExcelChartExNumericData data)
            {
                var node = data.TopNode;
                var innerXml = data.TopNode.InnerXml;
                node.ParentNode.RemoveChild(node);

                var newNode = CreateNode("cx:strDim", false, true);
                newNode.InnerXml = innerXml;
                var nd = new ExcelChartExStringData(_serie._chart.WorkSheet.Name, NameSpaceManager, newNode);
                nd.Type = type;
                _list[index] = nd;
            }
            else
            {
                ((ExcelChartExStringData)_list[index]).Type = type;
            }
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

        internal ExcelChartExData GetValueDimension()
        {
            foreach(var d in _list)
            {
                if(d is ExcelChartExStringData s)
                {
                    if (s.Type != eStringDataType.Category) return d;
                }
                else if(d is ExcelChartExNumericData n)
                {
                    return n;
                }
            }
            return null;
        }
    }
}
