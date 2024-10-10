/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/30/2024         EPPlus Software AB       Pivot Table Conditional Formatting - EPPlus 7.4
 *************************************************************************************************/
using System;
using System.Xml;
using OfficeOpenXml.Core;
namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A collection of pivot areas used for conditional formatting
    /// </summary>
    public class ExcelPivotTableAreaConditionalFormattingCollection : EPPlusReadOnlyList<ExcelPivotTableAreaConditionalFormatting>
    {
        readonly ExcelPivotTable _pt;
        readonly XmlNamespaceManager _nsm;
        readonly XmlHelper _helper;
        internal ExcelPivotTableAreaConditionalFormattingCollection(ExcelPivotTable pt, XmlNode topNode)
        {
            _pt = pt;
            _nsm = pt.NameSpaceManager;
            _helper = XmlHelperFactory.Create(_nsm, topNode);
            foreach(XmlNode node in topNode.ChildNodes)
            {
                if(node.NodeType==XmlNodeType.Element && node.LocalName == "pivotArea" )
                {
                    var area = new ExcelPivotTableAreaConditionalFormatting(_nsm, node, _pt);
                    _list.Add(area);
                }
            }
        }
        /// <summary>
        /// Adds a new area for the one or more data fields
        /// </summary>
        /// <param name="fields">The data field(s) where the conditional formatting should be applied. If no fields are supplied all the pivot tables data fields will be added to the area</param>
        /// <returns>The pivot area for the conditional formatting</returns>
        public ExcelPivotTableAreaConditionalFormatting Add(params ExcelPivotTableDataField[] fields)
        {
            var node = _helper.CreateNode("d:pivotArea", false, true);
            var area = new ExcelPivotTableAreaConditionalFormatting(_nsm, node, _pt);
            if (fields == null && fields.Length > 0)
            {
                foreach (var field in fields)
                {
                    area.Conditions.DataFields.Add(field);
                }
            }
            else
            {
                if (_pt.DataFields.Count == 0)
                {
                    throw (new InvalidOperationException("Can't add a conditional format to a pivot table with no data fields."));
                }
                foreach (var field in _pt.DataFields)
                {
                    area.Conditions.DataFields.Add(field);
                }
            }

            _list.Add(area);
            return area;
        }
        /// <summary>
        /// /// Removes the the <paramref name="item"/> from the collection 
        /// </summary>
        /// <param name="item">The item to remove.</param>
        public void Remove(ExcelPivotTableAreaConditionalFormatting item)
        {
            item.TopNode.ParentNode.RemoveChild(item.TopNode);
            _list.Remove(item);
        }
        /// <summary>
        /// Removes the <see cref="ExcelPivotTableAreaStyle"/> at the <paramref name="index"/>
        /// </summary>
        /// <param name="index">The zero-based index in the collction to remove</param>
        public void RemoveAt(int index)
        {
            var x = _list[index];
            Remove(x);
        }

    }
}
