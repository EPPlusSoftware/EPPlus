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
    public class ExcelChartExDataPointCollection : XmlHelper, IEnumerable<ExcelChartExDataPoint>
    {
        ExcelChart _chart;
        private readonly SortedDictionary<int, ExcelChartExDataPoint> _dic = new SortedDictionary<int, ExcelChartExDataPoint>();
        internal ExcelChartExDataPointCollection(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) : base(ns, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            foreach (XmlNode pointNode in TopNode.SelectNodes(ExcelChartExDataPoint.topNodePath, ns))
            {
                var item = new ExcelChartExDataPoint(chart, ns, pointNode);
                _dic.Add(item.Index, item);
            }
            _chart = chart;

        }
        /// <summary>
        /// Adds a new datapoint to the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns>The datapoint</returns>
        public ExcelChartExDataPoint Add(int index)
        {
            return AddDp(index, null);
        }
        internal ExcelChartExDataPoint AddDp(int idx, string uniqueId = null)
        {
            if (_dic.ContainsKey(idx))
            {
                throw (new ArgumentException($"Point with index {idx} already exists"));
            }
            var pos = GetItemBefore(idx);

            XmlElement element = CreateElement(pos, uniqueId);
            var dp = new ExcelChartExDataPoint(_chart, NameSpaceManager, element, idx);

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
        private XmlElement CreateElement(int idx, string uniqueId = "")
        {
            XmlElement pointElement;
            if (_dic.Count == 0)
                pointElement = (XmlElement)CreateNode(ExcelChartDataPoint.topNodePath);
            else
            {
                pointElement = TopNode.OwnerDocument.CreateElement("c", "dPt", ExcelPackage.schemaChart);
                if (_dic.ContainsKey(idx))
                {
                    _dic[idx].TopNode.ParentNode.InsertAfter(pointElement, _dic[idx].TopNode);
                }
                else
                {
                    var first = _dic.Values.First().TopNode;
                    first.ParentNode.InsertBefore(pointElement, first);
                }
            }
            if (!string.IsNullOrEmpty(uniqueId))
            {
                if (_chart.IsType3D())
                {
                    pointElement.InnerXml = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d contourW=\"25400\"><a:contourClr><a:schemeClr val=\"lt1\"/></a:contourClr></a:sp3d></c:spPr><c:extLst><c:ext xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" uri = \"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\"><c16:uniqueId val=\"{" + uniqueId + "}\"/></c:ext></c:extLst>";
                }
                else
                {
                    pointElement.InnerXml = "<c:extLst><c:ext xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\" uri = \"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\"><c16:uniqueId val=\"{" + uniqueId + "}\"/></c:ext></c:extLst>";
                }
            }
            return pointElement;
        }

        private int GetItemBefore(int index)
        {
            if (_dic.ContainsKey(index - 1))
            {
                return index - 1;
            }
            int retIx = -1;
            foreach (int ix in _dic.Keys.OrderBy(x => x))
            {
                if (index < ix)
                {
                    return retIx;
                }
                retIx = ix;
            }
            return retIx;
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