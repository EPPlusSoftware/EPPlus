/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  6/4/2022         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class CssExporterBase : AbstractHtmlExporter
    {
        public CssExporterBase(HtmlExportSettings settings, ExcelRangeBase range)
        {
            Settings = settings;
            Require.Argument(range).IsNotNull("range");
            _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();
            ConditionalFormattingQuad(range);

            if (range.Addresses == null)
            {
                AddRange(range);
            }
            else
            {
                foreach (var address in range.Addresses)
                {
                    AddRange(range.Worksheet.Cells[address.Address]);
                }
            }
        }
        //abstract internal void WriteAndClearCollection(CssRuleCollection collection, CssTrueWriter writer);

        public CssExporterBase(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
        {
            Settings = settings;
            Require.Argument(ranges).IsNotNull("ranges");
            _ranges = ranges;
            ConditionalFormattingQuad(ranges[0]);
            //foreach(var range in _ranges) 
            //{
            //    ConditionalFormattingQuad(range);
            //}
        }

        protected HtmlExportSettings Settings;
        protected EPPlusReadOnlyList<ExcelRangeBase> _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();
        //protected Dictionary<string, List<ExcelConditionalFormattingRule>> _cfAtAddresses;
        protected QuadTree<ExcelConditionalFormattingRule> _cfQuadTree = null;

        private void AddRange(ExcelRangeBase range)
        {
            if (range.IsFullColumn && range.IsFullRow)
            {
                _ranges.Add(new ExcelRangeBase(range.Worksheet, range.Worksheet.Dimension.Address));
            }
            else
            {
                _ranges.Add(range);
            }
        }

        private void ConditionalFormattingQuad(ExcelRangeBase range)
        {
            if(_cfQuadTree == null)
            {
                _cfQuadTree = new QuadTree<ExcelConditionalFormattingRule>(range);
            }

            foreach(ExcelConditionalFormattingRule rule in range.Worksheet.ConditionalFormatting)
            {
                _cfQuadTree.Add(new QuadRange(rule.Address), rule);
            }
        }
    }
}
