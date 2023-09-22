using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.HtmlExport.Parsers
{
    internal static class AttributeParser
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

        internal static string GetClassAttributeFromStyle(ExcelRangeBase cell, bool isHeader, HtmlExportSettings settings, 
            string additionalClasses, Dictionary<string, List<ExcelConditionalFormattingRule>> cfCollection, 
            Dictionary<string, int> _styleCache, Dictionary<string, int> _dxfStyleCache)
        {
            string cls = string.IsNullOrEmpty(additionalClasses) ? "" : additionalClasses;
            int styleId = cell.StyleID;
            ExcelStyles styles = cell.Worksheet.Workbook.Styles;

            if (styleId < 0 || styleId >= styles.CellXfs.Count)
            {
                return "";
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
                    return cls;
                return "";
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
            if (_styleCache.ContainsKey(key))
            {
                id = _styleCache[key];
            }
            else
            {
                id = _styleCache.Count + 1;
                _styleCache.Add(key, id);
            }

            cls += $" {styleClassPrefix}{settings.CellStyleClassName}{id}";

            if (cfCollection.ContainsKey(cell.Address))
            {
                int dxfId;
                string dxfKey;

                for (int i = 0; i < cfCollection[cell.Address].Count(); i++)
                {
                    if (cfCollection[cell.Address][i].ShouldApplyToCell(cell))
                    {
                        dxfKey = cfCollection[cell.Address][i].Style.Id;

                        if (_dxfStyleCache.ContainsKey(dxfKey))
                        {
                            dxfId = _dxfStyleCache[dxfKey];
                        }
                        else
                        {
                            dxfId = _dxfStyleCache.Count + 1;
                            _dxfStyleCache.Add(dxfKey, id);
                        }

                        cls += $" {styleClassPrefix}{settings.CellStyleClassName}dxf id{dxfId}";
                    }
                }
            }

            return cls.Trim();
           // AddAttribute("class", cls.Trim());
        }

    }
}
