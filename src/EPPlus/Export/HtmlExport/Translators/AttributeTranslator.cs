﻿/*************************************************************************************************
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
                                specials += "height: 100%";
                                cls += $" {styleClassPrefix}{settings.CellStyleClassName}-irrelevantTmp";
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

                                cls += $" {styleClassPrefix}{settings.CellStyleClassName}-dxf id{dxfId}";
                                break;
                        }
                    }
                }
            }

            return new List<string> { cls.Trim(), specials };
        }

        internal static List<HTMLElement> ConditionalFormattingsDatabarToHTML(ExcelRangeBase cell, HtmlExportSettings settings, 
            ExporterContext context, HTMLElement parentElement)
        {
            var dataBarElements = new List<HTMLElement>();

            var cfItems = context._cfQuadTree.GetIntersectingRangeItems
                (new QuadRange(new ExcelAddress(cell.Address)));

            for (int i = 0; i < cfItems.Count(); i++)
            {
                if (cfItems[i].Value.ShouldApplyToCell(cell))
                {
                    switch (cfItems[i].Value.Type)
                    {
                        case eExcelConditionalFormattingRuleType.DataBar:
                            var bar = (ExcelConditionalFormattingDataBar)cfItems[i].Value;

                            var dbParent = new HTMLElement("div");
                            dbParent.AddAttribute("class", $"{settings.StyleClassPrefix}pRelParent");

                            parentElement.AddChildElement(dbParent);

                            var divNeg = new HTMLElement("div");
                            var divPos = new HTMLElement("div");
                            var divContent = new HTMLElement("div");

                            var prefix = $"{settings.StyleClassPrefix}{settings.CellStyleClassName}";

                            if (Convert.ToDouble(cell.Value) < 0)
                            {
                                divNeg.AddAttribute("class", $"{settings.StyleClassPrefix}relChildLeft neg-dbar {prefix}-db-neg{bar.DxfId} leftWidth{bar.DxfId}");
                                divPos.AddAttribute("class", $"{settings.StyleClassPrefix}relChildRight");

                                divNeg.AddAttribute("style", bar.ApplyStyleOverride(cell));
                            }
                            else
                            {
                                divNeg.AddAttribute("class", $"{settings.StyleClassPrefix}relChildLeft leftWidth{bar.DxfId}");
                                divPos.AddAttribute("class", $"{settings.StyleClassPrefix}relChildRight pos-dbar {prefix}-db-pos{bar.DxfId}");
                                divPos.AddAttribute("style", bar.ApplyStyleOverride(cell));
                            }

                            if (cell.StyleID > 0)
                            {
                                string hAlign = GetHorizontalAlignmentDBar(cell.Style.HorizontalAlignment);
                                string vAlign = GetVerticalAlignmentDBar(cell.Style.VerticalAlignment);

                                divContent.AddAttribute("style", $"justify-content: {hAlign}; align-items: {vAlign};");
                            }

                            var dataType = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cell.Value);
                            AddDataFromDatabarCell(cell, dataType, settings, divContent);

                            divContent.AddAttribute("class", $"{settings.StyleClassPrefix}dbc");

                            dbParent.AddChildElement(divNeg);
                            dbParent.AddChildElement(divPos);
                            dbParent.AddChildElement(divContent);
                            break;
                    }
                }
            }
            return dataBarElements;
        }

        static void AddDataFromDatabarCell(ExcelRangeBase cell, string dataType, HtmlExportSettings settings, HTMLElement element)
        {
            if (dataType != ColumnDataTypeManager.HtmlDataTypes.String && settings.RenderDataAttributes)
            {
                var v = HtmlRawDataProvider.GetRawValue(cell.Value, dataType);
                if (string.IsNullOrEmpty(v) == false)
                {
                    element.AddAttribute($"data-{settings.DataValueAttributeName}", v);
                }
            }
            if (settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                element.AddAttribute("role", "cell");
            }

            if (cell.IsRichText)
            {
                element.Content = cell.RichText.HtmlText;
            }
            else
            {
                element.Content = ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, settings.Culture);
            }
        }

        static string GetVerticalAlignmentDBar(ExcelVerticalAlignment vAlign)
        {
            switch (vAlign)
            {
                case ExcelVerticalAlignment.Top:
                    return "top";
                case ExcelVerticalAlignment.Center:
                    return "center";
                case ExcelVerticalAlignment.Bottom:
                    return "end";
            }

            return "";
        }

        static string GetHorizontalAlignmentDBar(ExcelHorizontalAlignment hAlign)
        {

            switch (hAlign)
            {
                case ExcelHorizontalAlignment.Right:
                    return "right";
                case ExcelHorizontalAlignment.Center:
                case ExcelHorizontalAlignment.CenterContinuous:
                    return "center";
                case ExcelHorizontalAlignment.Left:
                    return "left";
                default:
                    return "right";
            }

            return "";
        }
    }
}
