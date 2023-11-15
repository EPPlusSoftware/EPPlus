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
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Export.HtmlExport.Parsers;
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
                    AccessibilitySettings accessibilitySettings, bool addRowScope, HtmlImage image, ExporterContext content)
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
            var classString = AttributeTranslator.GetClassAttributeFromStyle(cell, false, settings, imageCellClassName, content);

            if (!string.IsNullOrEmpty(classString))
            {
                writer.AddAttribute("class", classString);
            }

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

        public void Write(ExcelRangeBase cell, string dataType, HTMLElement element, HtmlExportSettings settings, 
            AccessibilitySettings accessibilitySettings, bool addRowScope, HtmlImage image, ExporterContext content)
        {
            if (dataType != ColumnDataTypeManager.HtmlDataTypes.String && settings.RenderDataAttributes)
            {
                var v = HtmlRawDataProvider.GetRawValue(cell.Value, dataType);
                if (string.IsNullOrEmpty(v) == false)
                {
                    element.AddAttribute($"data-{settings.DataValueAttributeName}", v);
                }
            }
            if (accessibilitySettings.TableSettings.AddAccessibilityAttributes)
            {
                element.AddAttribute("role", "cell");
                if (addRowScope)
                {
                    element.AddAttribute("scope", "row");
                }
            }
            var imageCellClassName = image == null ? "" : settings.StyleClassPrefix + "image-cell";
            var classString = AttributeTranslator.GetClassAttributeFromStyle(cell, false, settings, imageCellClassName, content);

            if (!string.IsNullOrEmpty(classString))
            {
                element.AddAttribute("class", classString);
            }

            //writer.RenderBeginTag(HtmlElements.TableData);
            HtmlExportImageUtil.AddImage(element, settings, image, cell.Value);
            if (cell.IsRichText)
            {
                element.Content = cell.RichText.HtmlText;
                //writer.Write(cell.RichText.HtmlText);
            }
            else
            {
                element.Content = ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, settings.Culture);

                //writer.Write(ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, settings.Culture));
            }

            //element.AddChildElement(element);
            //writer.RenderEndTag();
            //writer.ApplyFormat(settings.Minify);
        }
#if !NET35
        public async Task WriteAsync(ExcelRangeBase cell, string dataType, EpplusHtmlWriter writer, HtmlExportSettings settings, 
            AccessibilitySettings accessibilitySettings, bool addRowScope, HtmlImage image, ExporterContext content)
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
            
            var classString = AttributeTranslator.GetClassAttributeFromStyle(cell, false, settings, imageCellClassName, content);

            if (!string.IsNullOrEmpty(classString))
            {
                writer.AddAttribute("class", classString);
            }

            //writer.SetClassAttributeFromStyle(cell, false, settings, imageCellClassName, cfRules);
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