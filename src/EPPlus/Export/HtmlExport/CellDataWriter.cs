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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
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
        private readonly CompileResultFactory _compileResultFactory = new CompileResultFactory();
        public void Write(ExcelRangeBase cell, string dataType, EpplusHtmlWriter writer, HtmlExportSettings settings, bool addRowScope)
        {
            if (dataType != ColumnDataTypeManager.HtmlDataTypes.String)
            {
                var v = HtmlRawDataProvider.GetRawValue(cell.Value, dataType);
                if (string.IsNullOrEmpty(v)==false)
                {
                    writer.AddAttribute("data-value", v);
                }
            }
            if (settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "cell");
                if(addRowScope)
                {
                    writer.AddAttribute("scope", "row");
                }
            }
            writer.SetClassAttributeFromStyle(cell.StyleID, cell.Worksheet.Workbook.Styles);
            writer.RenderBeginTag(HtmlElements.TableData);
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
#if !NET35 && !NET40
        public async Task WriteAsync(ExcelRangeBase cell, string dataType, EpplusHtmlWriter writer, HtmlTableExportSettings settings, bool addRowScope)
        {
            if (dataType != ColumnDataTypeManager.HtmlDataTypes.String)
            {
                var v = HtmlRawDataProvider.GetRawValue(cell, dataType);
                if (string.IsNullOrEmpty(v) == false)
                {
                    writer.AddAttribute("data-value", v);
                }
            }
            if (settings.Accessibility.TableSettings.AddAccessibilityAttributes)
            {
                writer.AddAttribute("role", "cell");
                if (addRowScope)
                {
                    writer.AddAttribute("scope", "row");
                }
            }
            writer.SetClassAttributeFromStyle(cell.StyleID, cell.Worksheet.Workbook.Styles);
            await writer.RenderBeginTagAsync(HtmlElements.TableData);
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