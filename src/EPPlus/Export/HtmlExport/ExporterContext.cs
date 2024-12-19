/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Core.RangeQuadTree;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class ExporterContext
    {
        internal readonly StyleCache _styleCache = new StyleCache();
        internal readonly StyleCache _dxfStyleCache = new StyleCache();
        internal readonly StyleCacheAlt XfsGenericCache = new StyleCacheAlt();

        internal QuadTree<ExcelConditionalFormattingRule> _cfQuadTree = null;

        internal ExporterContext() 
        {
        }

        internal void InitializeQuadTree(ExcelRangeBase range)
        {
            if (_cfQuadTree == null)
            {
                _cfQuadTree = new QuadTree<ExcelConditionalFormattingRule>(range);
            }

            //TODO: only for relevant range not worksheet
            foreach (ExcelConditionalFormattingRule rule in range.Worksheet.ConditionalFormatting)
            {
                foreach (var address in rule.Address.GetAllAddresses())
                {
                    if (address.Collide(range) != ExcelAddressBase.eAddressCollition.No)
                    {
                        _cfQuadTree.Add(new QuadRange(address), rule);
                    }
                }
            }
        }
    }
}
