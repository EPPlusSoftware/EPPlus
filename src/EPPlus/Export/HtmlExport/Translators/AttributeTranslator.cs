﻿using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;

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

            if (styleId == 0 || HasStyle(xfs) == false)
            {
                if (string.IsNullOrEmpty(cls) == false)
                    return new List<string> { cls };
                return new List<string> { "" };
            }

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


            int dxfId;
            string dxfKey;

            string specials = "";
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
                            cls += $" {styleClassPrefix}{settings.CellStyleClassName}-irrelevantTmp";
                            //var bar = (ExcelConditionalFormattingDataBar)cfItems[i].Value;
                            //specials += $"{settings.StyleClassPrefix}{settings.CellStyleClassName}-databar-positive-1";
                            //if(bar.NegativeFillColor != null)
                            //{
                            //    specials += $"{settings.StyleClassPrefix}{settings.CellStyleClassName}-databar-negative-1";
                            //}
                            //specials += bar.ApplyStyleOverride(cell);
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
                                dxfStyleCache.Add(dxfKey, id);
                            }

                            cls += $" {styleClassPrefix}{settings.CellStyleClassName}-dxf id{dxfId}";
                            break;
                    }
                }
            }

            return new List<string> { cls.Trim(), specials };
        }

        internal static List<HTMLElement> ConditionalFormattingsDatabarToHTML(ExcelRangeBase cell, HtmlExportSettings settings, 
            ExporterContext context, HTMLElement parentElement, string parentClasses)
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
                            var flexParent = new HTMLElement("div");
                            flexParent.AddAttribute("class", $"{settings.StyleClassPrefix}fp");
                            flexParent.AddAttribute("style", $"height: 100%;");
                            parentElement.AddChildElement(flexParent);

                            if (bar.NegativeFillColor != null)
                            {
                                var divNeg = new HTMLElement("div");
                                var divPos = new HTMLElement("div");
                                if (Convert.ToDouble(cell.Value) < 0)
                                {
                                    divNeg.AddAttribute("class", $"{settings.StyleClassPrefix}fch {settings.StyleClassPrefix}bdr {settings.StyleClassPrefix}{settings.CellStyleClassName}-databar-negative-1");
                                    divPos.AddAttribute("class", $"{settings.StyleClassPrefix}fch");
                                    divPos.AddAttribute("style", $"align-self: end;");

                                    divNeg.AddAttribute("style", bar.ApplyStyleOverride(cell));
                                }
                                else
                                {
                                    divNeg.AddAttribute("class", $"{settings.StyleClassPrefix}fch");
                                    divPos.AddAttribute("class", $"{settings.StyleClassPrefix}fch {settings.StyleClassPrefix}bdr {settings.StyleClassPrefix}{settings.CellStyleClassName}-databar-positive-1");
                                    divPos.AddAttribute("style", bar.ApplyStyleOverride(cell)+";align-self: end;");
                                }
                                divPos.Content = cell.Value.ToString();

                                flexParent.AddChildElement(divNeg);
                                flexParent.AddChildElement(divPos);
                            }
                            else
                            {
                                parentClasses += bar.ApplyStyleOverride(cell);
                                parentClasses += $"{settings.StyleClassPrefix}{settings.CellStyleClassName}-databar-positive-1";
                            }
                        break;
                    }
                }
            }
            return dataBarElements;
        }

        //void SpecialOperation(ExcelConditionalFormattingRule rule, ExcelAddress address)
        //{
        //    switch (rule.Type) 
        //    {
        //        case eExcelConditionalFormattingRuleType.TwoColorScale:
        //        case eExcelConditionalFormattingRuleType.ThreeColorScale:

        //            var castType = (ExcelConditionalFormattingTwoColorScale)rule.As.TwoColorScale;
        //            castType.ApplyStyleOverride()
        //            break;
        //    }
        //}

    }
}
