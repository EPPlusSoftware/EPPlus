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
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Core;
using System;
using System.Linq;
using System.Xml;
namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A collection of pivot areas used for styling a pivot table.
    /// </summary>
    public class ExcelPivotTableConditionalFormattingCollection : EPPlusReadOnlyList<ExcelPivotTableConditionalFormatting>
    {
        ExcelConditionalFormattingCollection _conditionalFormatting;
        XmlHelper _xmlHelper;
        ExcelPivotTable _pt;
        internal ExcelPivotTableConditionalFormattingCollection(ExcelPivotTable pt)
        {
            _pt = pt;
            _conditionalFormatting = pt.WorkSheet.ConditionalFormatting;
            foreach (XmlNode node in pt.GetNodes("d:conditionalFormats/d:conditionalFormat"))
            {
                var cf = new ExcelPivotTableConditionalFormatting(_pt.NameSpaceManager, node, _pt);
                _list.Add(cf);
            }
        }
        /// <summary>
        /// Adds a conditional formatting pivot area for the pivot tables data field(cf).
        /// Note that only conditional formattings for data is support. Conditional formattings for Lables, data buttons and other pivot areas must be added using the <see cref="ExcelWorksheet.ConditionalFormatting" /> collection.
        /// </summary>
        /// <param name="ruleType">The type of conditional formatting rule</param>
        /// <param name="fields">The data field(cf) in the pivot table to apply the rule. If no data field is provided, all data field in the collection will be added to the area.The area will be added to the <see cref="ExcelPivotTableConditionalFormatting.Areas" /> collection</param>
        /// <returns>The rule</returns>
        public ExcelPivotTableConditionalFormatting Add(eExcelPivotTableConditionalFormattingRuleType ruleType, params ExcelPivotTableDataField[] fields)
        {
            var cfFormatNode = GetTopNode();
            var ct = new ExcelPivotTableConditionalFormatting(_pt.NameSpaceManager, cfFormatNode, _pt, (eExcelConditionalFormattingRuleType)ruleType);
            var a = ct.Areas.Add(fields);
            _list.Add(ct);
            return ct;
        }

        internal void Remove(ExcelPivotTableConditionalFormatting x)
        {
            x.TopNode.ParentNode.RemoveChild(x.TopNode);
            _pt.WorkSheet.ConditionalFormatting.Remove(x.ConditionalFormatting);
            _list.Remove(x);
        }
        internal void RemoveAt(int index)
        {
            var x = _list[index];
            Remove(x);
        }

        private XmlNode GetTopNode()
        {
            if (_xmlHelper == null)
            {
                var node = _pt.CreateNode("d:conditionalFormats");
                _xmlHelper = XmlHelperFactory.Create(_pt.NameSpaceManager, node);
            }
            
            var retNode = _xmlHelper.CreateNode("d:conditionalFormat", false,true);
            retNode.InnerXml = $"<pivotAreas xmlns=\"{ExcelPackage.schemaMain}\"/>";
            return retNode;
        }
    }
}
