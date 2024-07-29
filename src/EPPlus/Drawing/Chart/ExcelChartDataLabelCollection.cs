/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// A collection of individually formatted datalabels
    /// </summary>
    public class ExcelChartDataLabelCollection : XmlHelper, IEnumerable<ExcelChartDataLabelItem>
    {
        ExcelChart _chart;
        private readonly List<ExcelChartDataLabelItem> _list;
        internal ExcelChartDataLabelCollection(ExcelChart chart, XmlNamespaceManager ns, XmlNode topNode, string[] schemaNodeOrder) : base(ns, topNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _list = new List<ExcelChartDataLabelItem>();
            foreach (XmlNode dataLabelNode in TopNode.SelectNodes("c:dLbl", ns))
            {
                _list.Add(new ExcelChartDataLabelItem(chart, ns, dataLabelNode, "", schemaNodeOrder));
            }
            _chart = chart;
        }
        /// <summary>
        /// Adds a new chart label to the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns></returns>
        public ExcelChartDataLabelItem Add(int index)
        {
            if (_list.Count == 0)
            {
                return CreateDataLabel(index);
            }
            else
            {
                var ix = GetItemAfter(index);
                if (_list[ix].Index == index)
                {
                    throw (new ArgumentException($"Data label with index {index} already exists"));
                }
                return CreateDataLabel(index);
            }
        }

        private ExcelChartDataLabelItem CreateDataLabel(int idx)
        {
            XmlElement element = CreateElement(idx);
            var dl = new ExcelChartDataLabelItem(_chart, NameSpaceManager, element, "dLbl", SchemaNodeOrder) { Index=idx };

            if (idx < _list.Count)
            {
                _list.Insert(idx, dl);
            }
            else
            {
                _list.Add(dl);
            }

            return dl;
        }

        private XmlElement CreateElement(int idx)
        {
            XmlElement pointNode;
            if (idx < _list.Count)
            {
                pointNode = TopNode.OwnerDocument.CreateElement("c", "dLbl", @"http://schemas.openxmlformats.org/drawingml/2006/chart");
                TopNode.InsertBefore(pointNode, _list[idx].TopNode);
            }
            else
            {
                pointNode = (XmlElement)CreateNode("c:dLbl", false, true);
            }
            return pointNode;
        }

        private int GetItemAfter(int index)
        {
            for (var i = 0; i < _list.Count; i++)
            {
                if (index >= _list[i].Index)
                {
                    return i;
                }
            }
            return _list.Count;
        }
        /// <summary>
        /// Indexer for the collection
        /// </summary>
        /// <param name="index">The index</param>
        /// <returns></returns>
        public ExcelChartDataLabelItem this[int index]
        {
            get
            {
                return (_list[index]);
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
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelChartDataLabelItem> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
    }
}