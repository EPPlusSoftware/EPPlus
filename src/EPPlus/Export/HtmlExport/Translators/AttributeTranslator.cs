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
using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace OfficeOpenXml.Export.HtmlExport.Parsers
{
    internal static class AttributeTranslator
    {
        internal static string GetClassAttributeFromStyle(ExcelRangeBase cell, bool isHeader, HtmlExportSettings settings, 
            string additionalClasses, ExporterContext context)
        {
            string cls = string.IsNullOrEmpty(additionalClasses) ? "" : additionalClasses;
            int styleId = cell.StyleID;
            ExcelStyles styles = cell.Worksheet.Workbook.Styles;

            if (styleId < 0 || styleId >= styles.CellXfs.Count)
            {
                return  "";
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

            StyleXml xfsGeneric;
            if (context.XfsGenericCache.ContainsKey(styleId))
            {
                xfsGeneric = (StyleXml)context.XfsGenericCache[styleId];
            }
            else
            {
                xfsGeneric = new StyleXml(xfs);
                context.XfsGenericCache.Add(styleId, xfsGeneric);
            }

            if (styleId > 0 && xfsGeneric.HasStyle == true)
            {
                string key = xfsGeneric.GetStyleKey();

                var ma = cell.Worksheet.MergedCells[cell._fromRow, cell._fromCol];
                if (ma != null)
                {
                    var address = new ExcelAddressBase(ma);
                    var bottomStyleId = cell.Worksheet._values.GetValue(address._toRow, address._fromCol)._styleId;
                    var rightStyleId = cell.Worksheet._values.GetValue(address._fromRow, address._toCol)._styleId;
                    key += bottomStyleId + "|" + rightStyleId;
                }

                context._styleCache.GetOrCreateId(key, out int id);

                if (string.IsNullOrEmpty(cls))
                {
                    cls += $"{styleClassPrefix}{settings.CellStyleClassName}{id}";
                }
                else
                {
                    cls += $" {styleClassPrefix}{settings.CellStyleClassName}{id}";
                }
            }

            return cls;
        }

        internal static List<string> GetConditionalFormattings(ExcelRangeBase cell, HtmlExportSettings settings, ExporterContext context, ref string cls)
        {
            string inlineStyles = "";
            string extras = "";

            var styleClassPrefix = settings.StyleClassPrefix;
            var dxfStyleCache = context._dxfStyleCache;

            if (settings.RenderConditionalFormattings)
            {
                int dxfId;
                string dxfKey;

                var prefix = $" { styleClassPrefix }{ settings.DxfStyleClassName}";
                var middle = $"{settings.ConditionalFormattingClassName}-ic";
                var iconPrefix = $" {styleClassPrefix}{settings.IconPrefix}";

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
                                inlineStyles += ((ExcelConditionalFormattingTwoColorScale)cfItems[i].Value).ApplyStyleOverride(cell);
                                break;
                            case eExcelConditionalFormattingRuleType.ThreeColorScale:
                                inlineStyles += ((ExcelConditionalFormattingThreeColorScale)cfItems[i].Value).ApplyStyleOverride(cell);
                                break;
                            case eExcelConditionalFormattingRuleType.ThreeIconSet:
                                dxfStyleCache.GetOrCreateId(cfItems[i].Value.Uid, out dxfId);
                                var iconNameThree = GetIconName((ExcelConditionalFormattingThreeIconSet)cfItems[i].Value.As.ThreeIconSet, cell);
                                cls += $"{prefix}{dxfId}";
                                cls += AddIconClasses(iconNameThree, iconPrefix);
                                break;
                            case eExcelConditionalFormattingRuleType.FourIconSet:
                                dxfStyleCache.GetOrCreateId(cfItems[i].Value.Uid, out dxfId);
                                cls += $"{prefix}{dxfId}";
                                var iconNameFour = GetIconName((ExcelConditionalFormattingFourIconSet)cfItems[i].Value.As.FourIconSet, cell);
                                cls += AddIconClasses(iconNameFour, iconPrefix);
                                break;
                            case eExcelConditionalFormattingRuleType.FiveIconSet:
                                dxfStyleCache.GetOrCreateId(cfItems[i].Value.Uid, out dxfId);
                                cls += $"{prefix}{dxfId}";
                                var iconNameFive = GetIconName((ExcelConditionalFormattingFiveIconSet)cfItems[i].Value.As.FiveIconSet, cell);
                                cls += AddIconClasses(iconNameFive, iconPrefix);
                                break;
                            case eExcelConditionalFormattingRuleType.DataBar:
                                dxfStyleCache.GetOrCreateId(cfItems[i].Value.Uid, out dxfId);
                                cls += $" {styleClassPrefix}{settings.DatabarPrefix}-shared";
                                //var dbar = (ExcelConditionalFormattingDataBar)cfItems[i].Value;

                                var ruleName = $"{prefix}{dxfId}-";
                                var realValue = Convert.ToDouble(cell.Value);

                                //Color borderColor;
                                if (realValue > 0)
                                {
                                    cls += ruleName + "pos";
                                }
                                else
                                {
                                    cls += ruleName + "neg";
                                }

                                cls += $" {styleClassPrefix}{cell.Address}-{settings.DatabarPrefix}";

                                break;
                            default:
                                dxfKey = cfItems[i].Value.Style.Id;
                                dxfStyleCache.GetOrCreateId(dxfKey, out dxfId);

                                cls += $"{prefix}{dxfId}";
                                break;
                        }
                    }
                }
            }

            if (extras != "")
            {
                return new List<string> { inlineStyles, extras };
            }

            return new List<string> { inlineStyles };
        }

        internal static string GetIconName<T>(ExcelConditionalFormattingIconSetBase<T> set, ExcelRangeBase cell)
            where T : struct, Enum
        {
            var iconName = set.GetIconName(cell);
            return iconName;
        }

        internal static string AddIconClasses(string iconName, string prefix)
        {
        string retString = "";
            retString += $"{prefix}-shared";

            if(iconName != "")
            {
                retString += $"{prefix}-{iconName}";
            }

            return retString;
        }
    }
}
