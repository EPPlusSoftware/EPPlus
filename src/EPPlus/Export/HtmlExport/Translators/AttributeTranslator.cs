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
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.HtmlExport.Parsers
{
    internal static class AttributeTranslator
    {
        internal static bool HasStyle(ExcelXfs xfs)
        {
            return xfs.FontId > 0 ||
                   xfs.FillId > 0 ||
                   xfs.BorderId > 0 ||
                   xfs.HorizontalAlignment != ExcelHorizontalAlignment.General ||
                   xfs.VerticalAlignment != ExcelVerticalAlignment.Bottom ||
                   xfs.TextRotation != 0 ||
                   xfs.Indent > 0 ||
                   xfs.WrapText;
        }

        internal static string GetStyleKey(ExcelXfs xfs)
        {
            var fbfKey = ((ulong)(uint)xfs.FontId << 32 | (uint)xfs.BorderId << 16 | (uint)xfs.FillId);
            return fbfKey.ToString() + "|" + ((int)xfs.HorizontalAlignment).ToString() + "|" + ((int)xfs.VerticalAlignment).ToString() + "|" + xfs.Indent.ToString() + "|" + xfs.TextRotation.ToString() + "|" + (xfs.WrapText ? "1" : "0");
        }

        internal static List<string> GetClassAttributeFromStyle(ExcelRangeBase cell, bool isHeader, HtmlExportSettings settings, 
            string additionalClasses, ExporterContext context)
        {
            string cls = string.IsNullOrEmpty(additionalClasses) ? "" : additionalClasses;
            int styleId = cell.StyleID;
            ExcelStyles styles = cell.Worksheet.Workbook.Styles;

            var styleCache = context._styleCache;
            var dxfStyleCache = context._dxfStyleCache;

            if (styleId < 0 || styleId >= styles.CellXfs.Count)
            {
                return new List<string> { "" };
            }

            var xfs = styles.CellXfs[styleId];
            var styleClassPrefix = settings.StyleClassPrefix;
            if (settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.CellDataType &&
               xfs.HorizontalAlignment == ExcelHorizontalAlignment.General)
            {
                if (ConvertUtil.IsNumericOrDate(cell.Value))
                {
                    cls = $"{styleClassPrefix}ar";
                }
                else if (isHeader)
                {
                    cls = $"{styleClassPrefix}al";
                }
            }

            var returnList = new List<string> { "" };

            if (styleId == 0 || HasStyle(xfs) == false)
            {
                if (string.IsNullOrEmpty(cls) == false)
                    returnList[0] = cls;
            }
            else
            {
                string key = GetStyleKey(xfs);

                var ma = cell.Worksheet.MergedCells[cell._fromRow, cell._fromCol];
                if (ma != null)
                {
                    var address = new ExcelAddressBase(ma);
                    var bottomStyleId = cell.Worksheet._values.GetValue(address._toRow, address._fromCol)._styleId;
                    var rightStyleId = cell.Worksheet._values.GetValue(address._fromRow, address._toCol)._styleId;
                    key += bottomStyleId + "|" + rightStyleId;
                }

                int id;
                if (styleCache.ContainsKey(key))
                {
                    id = styleCache[key];
                }
                else
                {
                    id = styleCache.Count + 1;
                    styleCache.Add(key, id);
                }

                cls += $" {styleClassPrefix}{settings.CellStyleClassName}{id}";
            }

            string specials = "";

            if (settings.RenderConditionalFormattings)
            {
                int dxfId;
                string dxfKey;

                List<string> extraClasses = new List<string>();

                var cfItems = context._cfQuadTree.GetIntersectingRangeItems
                    (new QuadRange(new ExcelAddress(cell.Address)));

                for (int i = 0; i < cfItems.Count(); i++)
                {
                    if (cfItems[i].Value.ShouldApplyToCell(cell))
                    {
                        switch (cfItems[i].Value.Type)
                        {
                            case eExcelConditionalFormattingRuleType.TwoColorScale:
                                specials += ((ExcelConditionalFormattingTwoColorScale)cfItems[i].Value).ApplyStyleOverride(cell);
                                break;
                            case eExcelConditionalFormattingRuleType.ThreeColorScale:
                                specials += ((ExcelConditionalFormattingThreeColorScale)cfItems[i].Value).ApplyStyleOverride(cell);
                                break;
                            case eExcelConditionalFormattingRuleType.DataBar:
                                break;
                            default:
                                dxfKey = cfItems[i].Value.Style.Id;

                                if (dxfStyleCache.ContainsKey(dxfKey))
                                {
                                    dxfId = dxfStyleCache[dxfKey];
                                }
                                else
                                {
                                    dxfId = dxfStyleCache.Count + 1;
                                    dxfStyleCache.Add(dxfKey, dxfId);
                                }

                                cls += $" {styleClassPrefix}{settings.DxfStyleClassName}{dxfId}";
                                break;
                        }
                    }
                }
            }

            return new List<string> { cls.Trim(), specials };
        }
    }
}
