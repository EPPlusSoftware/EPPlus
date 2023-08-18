/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif
namespace OfficeOpenXml.Export.HtmlExport
{
    internal class CellDataWriter
    {
        public void Write(ExcelRangeBase cell, string dataType, EpplusHtmlWriter writer, HtmlExportSettings settings, 
            AccessibilitySettings accessibilitySettings, bool addRowScope, HtmlImage image, Dictionary<string, List<ExcelConditionalFormattingRule>> cfRules)
        {
            if (dataType != ColumnDataTypeManager.HtmlDataTypes.String && settings.RenderDataAttributes)
            {
                var v = HtmlRawDataProvider.GetRawValue(cell.Value, dataType);
                if (string.IsNullOrEmpty(v) == false)
                {
                    writer.AddAttribute($"data-{settings.DataValueAttributeName}", v);
                }
            }
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "cell");
                if (addRowScope)
                {
                    writer.AddAttribute("scope", "row");
                }
            }
            var imageCellClassName = image == null ? "" : settings.StyleClassPrefix + "image-cell";
            writer.SetClassAttributeFromStyle(cell, false, settings, imageCellClassName, cfRules);
            writer.RenderBeginTag(HtmlElements.TableData);
            HtmlExportImageUtil.AddImage(writer, settings, image, cell.Value);
            if (cell.IsRichText)
            {
                writer.Write(cell.RichText.HtmlText);
            }
            else
            {
                writer.Write(ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, settings.Culture));
            }

            writer.RenderEndTag();
            writer.ApplyFormat(settings.Minify);
        }
#if !NET35
        public async Task WriteAsync(ExcelRangeBase cell, string dataType, EpplusHtmlWriter writer, HtmlExportSettings settings, 
            AccessibilitySettings accessibilitySettings, bool addRowScope, HtmlImage image, Dictionary<string, List<ExcelConditionalFormattingRule>> cfRules)
        {
            if (dataType != ColumnDataTypeManager.HtmlDataTypes.String && settings.RenderDataAttributes)
            {
                var v = HtmlRawDataProvider.GetRawValue(cell.Value, dataType);
                if (string.IsNullOrEmpty(v) == false)
                {
                    writer.AddAttribute($"data-{settings.DataValueAttributeName}", v);
                }
            }
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "cell");
                if (addRowScope)
                {
                    writer.AddAttribute("scope", "row");
                }
            }
            var imageCellClassName = image == null ? "" : settings.StyleClassPrefix + "image-cell";
            writer.SetClassAttributeFromStyle(cell, false, settings, imageCellClassName, cfRules);
            await writer.RenderBeginTagAsync(HtmlElements.TableData);
            HtmlExportImageUtil.AddImage(writer, settings, image, cell.Value);
            if (cell.IsRichText)
            {
                await writer.WriteAsync(cell.RichText.HtmlText);
            }
            else
            {
                await writer.WriteAsync(ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, settings.Culture));
            }
            await writer.RenderEndTagAsync();
            await writer.ApplyFormatAsync(settings.Minify);
        }
#endif
    }
}