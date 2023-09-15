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

namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// A collection of filter columns for an autofilter of table in a worksheet
    /// </summary>
    public class ExcelFilterColumnCollection : XmlHelper, IEnumerable<ExcelFilterColumn>
    {
        SortedDictionary<int, ExcelFilterColumn> _columns = new SortedDictionary<int, ExcelFilterColumn>();
        ExcelAutoFilter _autoFilter;
        internal ExcelFilterColumnCollection(XmlNamespaceManager namespaceManager, XmlNode topNode, ExcelAutoFilter autofilter) : base(namespaceManager, topNode)
        {
            _autoFilter = autofilter;
            foreach (XmlElement node in topNode.SelectNodes("d:filterColumn", namespaceManager))
            {
                if(!int.TryParse(node.Attributes["colId"].Value, out int position))
                {
                    throw (new Exception("Invalid filter. Missing colId on filterColumn"));
                }
                switch (node.FirstChild?.Name)
                {
                    case "filters":
                        _columns.Add(position, new ExcelValueFilterColumn(namespaceManager, node));
                        break;
                    case "customFilters":
                        _columns.Add(position, new ExcelCustomFilterColumn(namespaceManager, node));
                        break;
                    case "colorFilter":
                        _columns.Add(position, new ExcelColorFilterColumn(namespaceManager, node));
                        break;
                    case "iconFilter":
                        _columns.Add(position, new ExcelIconFilterColumn(namespaceManager, node));
                        break;
                    case "dynamicFilter":
                        _columns.Add(position, new ExcelDynamicFilterColumn(namespaceManager, node));
                        break;
                    case "top10":
                        _columns.Add(position, new ExcelTop10FilterColumn(namespaceManager, node));
                        break;
                }
            }
        }
        /// <summary>
        /// Number of items in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _columns.Count;
            }
        }
        internal XmlNode Add(int position, string topNodeName)
        {
            XmlElement node;
            if (position >= _autoFilter.Address.Columns)
            {
                throw (new ArgumentOutOfRangeException("Position is outside of the range"));
            }
            if (_columns.ContainsKey(position))
            {
                throw (new ArgumentOutOfRangeException("Position already exists"));
            }
            foreach (var c in _columns.Values)
            {
                if (c.Position > position)
                {
                    node = GetColumnNode(position, topNodeName);
                    return c.TopNode.ParentNode.InsertBefore(node, c.TopNode);
                }
            }
            node = GetColumnNode(position, topNodeName);
            return TopNode.AppendChild(node);
        }

        private XmlElement GetColumnNode(int position, string topNodeName)
        {
            XmlElement node = TopNode.OwnerDocument.CreateElement("filterColumn", ExcelPackage.schemaMain);
            node.SetAttribute("colId", position.ToString());
            var subNode = TopNode.OwnerDocument.CreateElement(topNodeName, ExcelPackage.schemaMain);
            node.AppendChild(subNode);
            return node;
        }
        /// <summary>
        /// Indexer of filtercolumns
        /// </summary>
        /// <param name="index">The column index starting from zero</param>
        /// <returns>A filter column</returns>
        public ExcelFilterColumn this[int index]
        {
            get
            {
                if(_columns.ContainsKey(index))
                {
                    return _columns[index];
                }
                else
                {
                    return null;
                }
            }
        }
        /// <summary>
        /// Adds a value filter for the specified column position
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The value filter</returns>
        public ExcelValueFilterColumn AddValueFilterColumn(int position)
        {
            var node = Add(position, "filters");
            var col = new ExcelValueFilterColumn(NameSpaceManager, node);
            _columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Adds a custom filter for the specified column position
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The custom filter</returns>
        public ExcelCustomFilterColumn AddCustomFilterColumn(int position)
        {
            var node = Add(position, "customFilters");
            var col= new ExcelCustomFilterColumn(NameSpaceManager, node);
            _columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Adds a color filter for the specified column position
        /// Note: EPPlus doesn't filter color filters when <c>ApplyFilter</c> is called.
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The color filter</returns>
        public ExcelColorFilterColumn AddColorFilterColumn(int position)
        {
            var node = Add(position, "colorFilter");
            var col = new ExcelColorFilterColumn(NameSpaceManager, node);
            _columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Adds a icon filter for the specified column position
        /// Note: EPPlus doesn't filter icon filters when <c>ApplyFilter</c> is called.
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The color filter</returns>
        public ExcelIconFilterColumn AddIconFilterColumn(int position)
        {
            var node = Add(position, "iconFilter");
            var col = new ExcelIconFilterColumn(NameSpaceManager, node);
            _columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Adds a top10 filter for the specified column position
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The top 10 filter</returns>
        public ExcelTop10FilterColumn AddTop10FilterColumn(int position)
        {
            var node = Add(position, "top10");
            var col = new ExcelTop10FilterColumn(NameSpaceManager, node);
            _columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Adds a dynamic filter for the specified column position
        /// </summary>
        /// <param name="position">The column position</param>
        /// <returns>The dynamic filter</returns>
        public ExcelDynamicFilterColumn AddDynamicFilterColumn(int position)
        {
            var node = Add(position, "dynamicFilter");
            var col = new ExcelDynamicFilterColumn(NameSpaceManager, node);
            _columns.Add(position, col);
            return col;
        }
        /// <summary>
        /// Gets the enumerator of the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        public IEnumerator<ExcelFilterColumn> GetEnumerator()
        {
            return _columns.Values.GetEnumerator();
        }
        /// <summary>
        /// Gets the enumerator for the collection
        /// </summary>
        /// <returns>The enumerator</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _columns.Values.GetEnumerator();
        }
        /// <summary>
        /// Remove the filter column with the position from the collection
        /// </summary>
        /// <param name="position">The index of the column to remove</param>
        public void RemoveAt(int position)
        {
            if(!_columns.ContainsKey(position))
            {
                throw new InvalidOperationException($"Column with position {position} does not exist in the filter collection");
            }
            Remove(_columns[position]);
        }
        /// <summary>
        /// Remove the filter column from the collection
        /// </summary>
        /// <param name="column">The column</param>
        public void Remove(ExcelFilterColumn column)
        {
            var node = column.TopNode;
            node.ParentNode.RemoveChild(node);
            _columns.Remove(column.Position);
        }
        /// <summary>
        /// Clear all columns
        /// </summary>
        public void Clear()
        {
            _columns.Clear();
        }

    }
}