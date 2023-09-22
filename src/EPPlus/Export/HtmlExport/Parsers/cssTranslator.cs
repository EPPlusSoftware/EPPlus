using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Parsers
{
    internal class cssTranslator
    {
        //internal async Task AddToCssAsyncCF(ExcelDxfStyleConditionalFormatting dxfs, string styleClassPrefix, string cellStyleClassName, int priorityID, string uid)
        //{
        //    if (dxfs != null)
        //    {
        //        if (IsAddedToCache(dxfs, out int id) || _addedToCssCf.Contains(id) == false)
        //        {
        //            _addedToCssCf.Add(id);
        //            await WriteClassAsync($".{styleClassPrefix}{cellStyleClassName}-dxf-{id}{{", _settings.Minify);

        //            if (dxfs.Fill != null)
        //            {
        //                await WriteFillStylesAsync(dxfs.Fill);
        //            }

        //            if (dxfs.Font != null)
        //            {
        //                await WriteFontStylesAsync(dxfs.Font);
        //            }

        //            if (dxfs.Border != null)
        //            {
        //                await WriteBorderStylesAsync(dxfs.Border.Top, dxfs.Border.Bottom, dxfs.Border.Left, dxfs.Border.Right);
        //            }

        //            await WriteClassEndAsync(_settings.Minify);
        //        }
        //    }
        //}

        //private async Task WriteFillStylesAsync(ExcelDxfFill f)
        //{
        //    if (_cssExclude.Fill) return;

        //    if (f.PatternType == ExcelFillStyle.Solid || f.PatternType == null)
        //    {
        //        if (f.BackgroundColor.Color != null)
        //        {
        //            await WriteCssItemAsync($"background-color:{GetDxfColor(f.BackgroundColor)};", _settings.Minify);
        //        }
        //    }
        //}

        //private async Task WriteFontStylesAsync(ExcelDxfFontBase f)
        //{

        //    bool hasDecoration = false;

        //    if (f.Color.HasValue && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Color))
        //    {
        //        await WriteCssItemAsync($"color:{GetDxfColor(f.Color)};", _settings.Minify);
        //    }
        //    if (f.Bold.HasValue && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Bold))
        //    {
        //        await WriteCssItemAsync("font-weight:bolder;", _settings.Minify);
        //    }
        //    if (f.Italic.HasValue && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Italic))
        //    {
        //        await WriteCssItemAsync("font-style:italic;", _settings.Minify);
        //    }
        //    if (f.Strike.HasValue && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Strike))
        //    {
        //        await WriteCssItemAsync("text-decoration:line-through", _settings.Minify);
        //        hasDecoration = true;
        //    }
        //    if (f.Underline.HasValue && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Underline))
        //    {
        //        if (!hasDecoration)
        //        {
        //            await WriteCssItemAsync("text-decoration:", _settings.Minify);
        //        }

        //        switch (f.Underline.Value)
        //        {
        //            case ExcelUnderLineType.Double:
        //            case ExcelUnderLineType.DoubleAccounting:
        //                await WriteCssItemAsync(" underline double;", _settings.Minify);
        //                break;
        //            default:
        //                await WriteCssItemAsync(" underline;", _settings.Minify);
        //                break;
        //        }
        //    }
        //    else if (hasDecoration)
        //    {
        //        await WriteCssItemAsync(";", _settings.Minify);
        //    }
        //}

        //private async Task WriteBorderStylesAsync(ExcelDxfBorderItem top, ExcelDxfBorderItem bottom, ExcelDxfBorderItem left, ExcelDxfBorderItem right)
        //{
        //    if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Top)) await WriteBorderItemAsync(top, "top");
        //    if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Bottom)) await WriteBorderItemAsync(bottom, "bottom");
        //    if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Left)) await WriteBorderItemAsync(left, "left");
        //    if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Right)) await WriteBorderItemAsync(right, "right");
        //}
    }
}
