/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/17/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style.XmlAccess;
using System.IO;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System;
using OfficeOpenXml.Utils;
using System.Text;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Interfaces;
using System.Linq;
using OfficeOpenXml.Export.HtmlExport.Exporters;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
#if !NET35 && !NET40
    internal partial class EpplusCssWriter : HtmlWriterBase
    {
        internal async Task RenderAdditionalAndFontCssAsync(string tableClass)
        {
            if (_cssSettings.IncludeSharedClasses == false) return;
            await WriteClassAsync($"table.{tableClass}{{", _settings.Minify);
            if (_cssSettings.IncludeNormalFont)
            {
                var ns = _ranges.First().Worksheet.Workbook.Styles.GetNormalStyle();
                if (ns != null)
                {
                    await WriteCssItemAsync($"font-family:{ns.Style.Font.Name};", _settings.Minify);
                    await WriteCssItemAsync($"font-size:{ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", _settings.Minify);
                }
            }
            foreach (var item in _cssSettings.AdditionalCssElements)
            {
                await WriteCssItemAsync($"{item.Key}:{item.Value};", _settings.Minify);
            }
            await WriteClassEndAsync(_settings.Minify);

            //Class for hidden rows.
            await WriteClassAsync($".{_settings.StyleClassPrefix}hidden {{", _settings.Minify);
            await WriteCssItemAsync($"display:none;", _settings.Minify);
            await WriteClassEndAsync(_settings.Minify);

            await WriteClassAsync($".{_settings.StyleClassPrefix}al {{", _settings.Minify);
            await WriteCssItemAsync($"text-align:left;", _settings.Minify);
            await WriteClassEndAsync(_settings.Minify);
            await WriteClassAsync($".{_settings.StyleClassPrefix}ar {{", _settings.Minify);
            await WriteCssItemAsync($"text-align:right;", _settings.Minify);
            await WriteClassEndAsync(_settings.Minify);

            var worksheets = _ranges.Select(x => x.Worksheet).Distinct().ToList();
            foreach (var ws in worksheets)
            {
                var clsName = HtmlExportTableUtil.GetWorksheetClassName(_settings.StyleClassPrefix, "dcw", ws, worksheets.Count > 1);
                await WriteClassAsync($".{clsName} {{", _settings.Minify);
                await WriteCssItemAsync($"width:{ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), ws.Workbook.MaxFontWidth)}px;", _settings.Minify);
                await WriteClassEndAsync(_settings.Minify);

                clsName = HtmlExportTableUtil.GetWorksheetClassName(_settings.StyleClassPrefix, "drh", ws, worksheets.Count > 1);
                await WriteClassAsync($".{clsName} {{", _settings.Minify);
                await WriteCssItemAsync($"height:{(int)(ws.DefaultRowHeight / 0.75)}px;", _settings.Minify);
                await WriteClassEndAsync(_settings.Minify);
            }

            if (_settings.Pictures.Include != ePictureInclude.Exclude && _settings.Pictures.CssExclude.Alignment == false)
            {
                await WriteClassAsync($"td.{_settings.StyleClassPrefix}image-cell {{", _settings.Minify);
                if (_settings.Pictures.AddMarginTop)
                {
                    await WriteCssItemAsync($"vertical-align:top;", _settings.Minify);
                }
                else
                {
                    await WriteCssItemAsync($"vertical-align:middle;", _settings.Minify);
                }
                if (_settings.Pictures.AddMarginTop)
                {
                    await WriteCssItemAsync($"text-align:left;", _settings.Minify);
                }
                else
                {
                    await WriteCssItemAsync($"text-align:center;", _settings.Minify);
                }
                await WriteClassEndAsync(_settings.Minify);
            }
        }
        internal async Task AddPictureToCssAsync(HtmlImage p)
        {
            var img = p.Picture.Image;
            string encodedImage;
            ePictureType? type;
            if (img.Type == ePictureType.Emz || img.Type == ePictureType.Wmz)
            {

                encodedImage = Convert.ToBase64String(ImageReader.ExtractImage(img.ImageBytes, out type));
            }
            else
            {
                encodedImage = Convert.ToBase64String(img.ImageBytes);
                type = img.Type.Value;
            }
            if (type == null) return;
            var pc = (IPictureContainer)p.Picture;
            if (_images.Contains(pc.ImageHash) == false)
            {
                string imageFileName = HtmlExportImageUtil.GetPictureName(p);
                await WriteClassAsync($"img.{_settings.StyleClassPrefix}image-{imageFileName}{{", _settings.Minify);
                await WriteCssItemAsync($"content:url('data:{GetContentType(type.Value)};base64,{encodedImage}');", _settings.Minify);
                if (_settings.Pictures.Position != ePicturePosition.DontSet)
                {
                    await WriteCssItemAsync($"position:{_settings.Pictures.Position.ToString().ToLower()};", _settings.Minify);
                }

                if (p.FromColumnOff != 0 && _settings.Pictures.AddMarginLeft)
                {
                    var leftOffset = p.FromColumnOff / ExcelPicture.EMU_PER_PIXEL;
                    await WriteCssItemAsync($"margin-left:{leftOffset}px;", _settings.Minify);
                }

                if (p.FromRowOff != 0 && _settings.Pictures.AddMarginTop)
                {
                    var topOffset = p.FromRowOff / ExcelPicture.EMU_PER_PIXEL;
                    await WriteCssItemAsync($"margin-top:{topOffset}px;", _settings.Minify);
                }


                await WriteClassEndAsync(_settings.Minify);
                _images.Add(pc.ImageHash);
            }
            await AddPicturePropertiesToCssAsync(p);
        }

        private async Task AddPicturePropertiesToCssAsync(HtmlImage image)
        {
            string imageName = HtmlExportTableUtil.GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);
            var width = image.Picture.GetPixelWidth();
            var height = image.Picture.GetPixelHeight();

            await WriteClassAsync($"img.{_settings.StyleClassPrefix}image-prop-{imageName}{{", _settings.Minify);
            if (_settings.Pictures.KeepOriginalSize == false)
            {
                if (width != image.Picture.Image.Bounds.Width)
                {
                    await WriteCssItemAsync($"max-width:{width:F0}px;", _settings.Minify);
                }
                if (height != image.Picture.Image.Bounds.Height)
                {
                    await WriteCssItemAsync($"max-height:{height:F0}px;", _settings.Minify);
                }
            }

            if (image.Picture.Border.LineStyle != null && _settings.Pictures.CssExclude.Border==false)
            {
                var border = GetDrawingBorder(image.Picture);
                await WriteCssItemAsync($"border:{border};", _settings.Minify);
            }
            await WriteClassEndAsync(_settings.Minify);
        }

        internal async Task AddToCssAsync(ExcelStyles styles, int styleId, string styleClassPrefix, string cellStyleClassName)
        {
            var xfs = styles.CellXfs[styleId];
            if (HasStyle(xfs))
            {
                if (IsAddedToCache(xfs, out int id)== false || _addedToCss.Contains(id) == false)
                {
                    _addedToCss.Add(id);
                    await WriteClassAsync($".{styleClassPrefix}{cellStyleClassName}{id}{{", _settings.Minify);
                    if (xfs.FillId > 0)
                    {
                        await WriteFillStylesAsync(xfs.Fill);
                    }
                    if (xfs.FontId > 0)
                    {
                        var ns = styles.GetNormalStyle();
                        await WriteFontStylesAsync(xfs.Font, ns.Style.Font);
                    }
                    if (xfs.BorderId > 0)
                    {
                        await WriteBorderStylesAsync(xfs.Border.Top, xfs.Border.Bottom, xfs.Border.Left, xfs.Border.Right);
                    }
                    await WriteStylesAsync(xfs);
                    await WriteClassEndAsync(_settings.Minify);
                }
            }
        }

        internal async Task AddToCssAsync(ExcelStyles styles, int styleId, int bottomStyleId, int rightStyleId, string styleClassPrefix, string cellStyleClassName)
        {
            var xfs = styles.CellXfs[styleId];
            var bXfs = styles.CellXfs[bottomStyleId];
            var rXfs = styles.CellXfs[rightStyleId];
            if (HasStyle(xfs) || bXfs.BorderId > 0 || rXfs.BorderId > 0)
            {
                if (IsAddedToCache(xfs, out int id, bottomStyleId, rightStyleId) == false || _addedToCss.Contains(id) == false)
                {
                    _addedToCss.Add(id);
                    await WriteClassAsync($".{styleClassPrefix}{cellStyleClassName}{id}{{", _settings.Minify);
                    if (xfs.FillId > 0)
                    {
                        WriteFillStyles(xfs.Fill);
                    }
                    if (xfs.FontId > 0)
                    {
                        var ns = styles.GetNormalStyle();
                        await WriteFontStylesAsync(xfs.Font, ns.Style.Font);
                    }
                    if (xfs.BorderId > 0 || bXfs.BorderId > 0 || rXfs.BorderId > 0)
                    {
                        await WriteBorderStylesAsync(xfs.Border.Top, bXfs.Border.Bottom, xfs.Border.Left, rXfs.Border.Right);
                    }
                    await WriteStylesAsync(xfs);
                    await WriteClassEndAsync(_settings.Minify);
                }
            }
        }

        private async Task WriteStylesAsync (ExcelXfs xfs)
        {
            if (_cssExclude.WrapText == false)
            {
                if (xfs.WrapText)
                {
                    await WriteCssItemAsync("white-space: break-spaces;", _settings.Minify);
                }
                else
                {
                    await WriteCssItemAsync("white-space: nowrap;", _settings.Minify);
                }
            }

            if (xfs.HorizontalAlignment != ExcelHorizontalAlignment.General && _cssExclude.HorizontalAlignment == false)
            {
                var hAlign = GetHorizontalAlignment(xfs);
                await WriteCssItemAsync($"text-align:{hAlign};", _settings.Minify);
            }

            if (xfs.VerticalAlignment != ExcelVerticalAlignment.Bottom && _cssExclude.VerticalAlignment == false)
            {
                var vAlign = GetVerticalAlignment(xfs);
                await WriteCssItemAsync($"vertical-align:{vAlign};", _settings.Minify);
            }
            if (xfs.TextRotation != 0 && _cssExclude.TextRotation == false)
            {
                if (xfs.TextRotation == 255)
                {
                    await WriteCssItemAsync($"writing-mode:vertical-lr;;", _settings.Minify);
                    await WriteCssItemAsync($"text-orientation:upright;", _settings.Minify);
                }
                else
                {
                    if (xfs.TextRotation > 90)
                    {
                        await WriteCssItemAsync($"transform:rotate({xfs.TextRotation - 90}deg);", _settings.Minify);
                    }
                    else
                    {
                        await WriteCssItemAsync($"transform:rotate({360 - xfs.TextRotation}deg);", _settings.Minify);
                    }
                }
            }

            if (xfs.Indent > 0 && _cssExclude.Indent == false)
            {
                await WriteCssItemAsync($"padding-left:{xfs.Indent * _cssSettings.IndentValue}{_cssSettings.IndentUnit};", _settings.Minify);
            }
        }

        private async Task WriteBorderStylesAsync(ExcelBorderItemXml top, ExcelBorderItemXml bottom, ExcelBorderItemXml left, ExcelBorderItemXml right)
        {
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Top)) await WriteBorderItemAsync(top, "top");
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Bottom)) await WriteBorderItemAsync(bottom, "bottom");
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Left)) await WriteBorderItemAsync(left, "left");
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Right)) await WriteBorderItemAsync(right, "right");
            //TODO add Diagonal
            //WriteBorderItem(b.DiagonalDown, "right");
            //WriteBorderItem(b.DiagonalUp, "right");
        }

        private async Task WriteBorderItemAsync(ExcelBorderItemXml bi, string suffix)
        {
            if (bi.Style != ExcelBorderStyle.None)
            {
                var sb = new StringBuilder();
                sb.Append(GetBorderItemLine(bi.Style, suffix));
                if (bi.Color!=null && bi.Color.Exists)
                {
                    sb.Append($" {GetColor(bi.Color)}");
                }
                sb.Append(";");

                await WriteCssItemAsync(sb.ToString(), _settings.Minify);
            }
        }

        private async Task WriteFontStylesAsync(ExcelFontXml f, ExcelFont nf)
        {
            if (string.IsNullOrEmpty(f.Name) == false && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Name) && f.Name.Equals(nf.Name) == false)
            {
                await WriteCssItemAsync($"font-family:{f.Name};", _settings.Minify);
            }
            if (f.Size > 0 && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Size) && f.Size != nf.Size)
            {
                await WriteCssItemAsync($"font-size:{f.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", _settings.Minify);
            }
            if (f.Color != null && f.Color.Exists && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Color) && AreColorEqual(f.Color, nf.Color) == false)
            {
                await WriteCssItemAsync($"color:{GetColor(f.Color)};", _settings.Minify);
            }
            if (f.Bold && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Bold) && nf.Bold != f.Bold)
            {
                await WriteCssItemAsync("font-weight:bolder;", _settings.Minify);
            }
            if (f.Italic && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Italic) && nf.Italic != f.Italic)
            {
                await WriteCssItemAsync("font-style:italic;", _settings.Minify);
            }
            if (f.Strike && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Strike) && nf.Strike != f.Strike)
            {
                await WriteCssItemAsync("text-decoration:line-through solid;", _settings.Minify);
            }
            if (f.UnderLineType != ExcelUnderLineType.None && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Underline) && f.UnderLineType != nf.UnderLineType)
            {
                switch (f.UnderLineType)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        await WriteCssItemAsync("text-decoration:underline double;", _settings.Minify);
                        break;
                    default:
                        await WriteCssItemAsync("text-decoration:underline solid;", _settings.Minify);
                        break;
                }
            }
        }

        private async Task WriteFillStylesAsync(ExcelFillXml f)
        {
            if (_cssExclude.Fill) return;
            if (f is ExcelGradientFillXml gf && gf.Type!=ExcelFillGradientType.None)
            {
                await WriteGradientAsync(gf);
            }
            else
            {
                if (f.PatternType == ExcelFillStyle.Solid)
                {
                    await WriteCssItemAsync($"background-color:{GetColor(f.BackgroundColor)};", _settings.Minify);
                }
                else
                {
                    await WriteCssItemAsync($"{PatternFills.GetPatternSvg(f.PatternType, GetColor(f.BackgroundColor), GetColor(f.PatternColor))}", _settings.Minify);
                }
            }
        }

        private async Task WriteGradientAsync(ExcelGradientFillXml gradient)
        {
            if (gradient.Type == ExcelFillGradientType.Linear)
            {
                await _writer.WriteAsync($"background: linear-gradient({(gradient.Degree + 90) % 360}deg");
            }
            else
            {
                await _writer.WriteAsync($"background:radial-gradient(ellipse {gradient.Right * 100}% {gradient.Bottom * 100}%");
            }

            await _writer.WriteAsync($",{GetColor(gradient.GradientColor1)} 0%");
            await _writer.WriteAsync($",{GetColor(gradient.GradientColor2)} 100%");

            await _writer.WriteAsync(");");
        }
        public async Task FlushStreamAsync()
        {
            await _writer.FlushAsync();
        }
    }
#endif
}
