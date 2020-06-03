/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart.ChartEx
{
    /// <summary>
    /// A collection of datapoints for a chart
    /// </summary>
    public class ExcelChartExDataPointCollection : XmlHelper, IEnumerable<ExcelChartExDataPoint>
    {
        ExcelChartExSerie _serie;
        internal readonly SortedDictionary<int, ExcelChartExDataPoint> _dic = new SortedDictionary<int, ExcelChartExDataPoint>();
        internal ExcelChartExDataPointCollection(ExcelChartExSerie serie, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) : base(ns, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            foreach (XmlNode pointNode in TopNode.SelectNodes(ExcelChartExDataPoint.dataPtPath, ns))
            {
                var item = new ExcelChartExDataPoint(serie, ns, pointNode, SchemaNodeOrder);
                _dic.Add(item.Index, item);
            }
            foreach (XmlElement stNode in TopNode.SelectNodes(ExcelChartExDataPoint.SubTotalPath, ns))
            {
                var ix = int.Parse(stNode.GetAttribute("val"));
                if(_dic.ContainsKey(ix))
                {
                    _dic[ix].SubTotal = true;
                }
                else
                {
                    var item = new ExcelChartExDataPoint(serie, ns, TopNode, SchemaNodeOrder);
                    _dic.Add(item.Index, item);
                }
            }
            _serie = serie;

        }
        /// <summary>
        /// Adds a new datapoint to the collection
        /// </summary>
        /// <param name="index">The zero based index</param>
        /// <returns>The datapoint</returns>
        public ExcelChartExDataPoint Add(int index)
        {
            return AddDp(index);
        }
        internal ExcelChartExDataPoint AddDp(int idx)
        {
            if (_dic.ContainsKey(idx))
            {
                throw (new ArgumentException($"Data point with index {idx} already exists"));
            }
            
            var dp = new ExcelChartExDataPoint(_serie, NameSpaceManager, TopNode, idx, SchemaNodeOrder);

            _dic.Add(idx, dp);

            return dp;
        }

        /// <summary>
        /// Checkes if the index exists in the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns>true if exists</returns>
        public bool ContainsKey(int index)
        {
            return _dic.ContainsKey(index);
        }
        /// <summary>
        /// Indexer
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns></returns>
        public ExcelChartExDataPoint this[int index]
        {
            get
            {
                return (_dic[index]);
            }
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _dic.Count;
            }
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelChartExDataPoint> GetEnumerator()
        {
            return _dic.Values.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _dic.Values.GetEnumerator();
        }
    }
}