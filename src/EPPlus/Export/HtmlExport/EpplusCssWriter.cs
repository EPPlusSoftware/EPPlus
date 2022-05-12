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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Style.XmlAccess;
using System.Collections.Generic;
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

namespace OfficeOpenXml.Export.HtmlExport
{
    internal partial class EpplusCssWriter : HtmlWriterBase
    {
        protected HtmlExportSettings _settings;
        protected CssExportSettings _cssSettings;
        protected CssExclude _cssExclude;
        List<ExcelRangeBase> _ranges;
        ExcelWorkbook _wb;
        ExcelTheme _theme;
        internal eFontExclude _fontExclude;
        internal eBorderExclude _borderExclude;
        internal HashSet<int> _addedToCss=new HashSet<int>();
        internal EpplusCssWriter(StreamWriter writer, List<ExcelRangeBase> ranges, HtmlExportSettings settings, CssExportSettings cssSettings, CssExclude cssExclude, Dictionary<string, int> styleCache) : base(writer, styleCache) 
        {
            _settings = settings;
            _cssSettings = cssSettings;
            _cssExclude = cssExclude;
            Init(ranges);
        }
        internal EpplusCssWriter(Stream stream, List<ExcelRangeBase> ranges, HtmlExportSettings settings, CssExportSettings cssSettings, CssExclude cssExclude, Dictionary<string, int> styleCache) : base(stream, settings.Encoding, styleCache)
        {
            _settings = settings;
            _cssSettings = cssSettings;
            _cssExclude = cssExclude;
            Init(ranges);
        }
        private void Init(List<ExcelRangeBase> ranges)
        {
            _ranges = ranges;
            _wb = _ranges[0].Worksheet.Workbook;
            if (_wb.ThemeManager.CurrentTheme == null)
            {
                _wb.ThemeManager.CreateDefaultTheme();
            }
            _theme = _wb.ThemeManager.CurrentTheme;
            _borderExclude = _cssExclude.Border;
            _fontExclude = _cssExclude.Font;
        }

        internal void RenderAdditionalAndFontCss(string tableClass)
        {
            if (_cssSettings.IncludeSharedClasses == false) return;
            WriteClass($"table.{tableClass}{{", _settings.Minify);
            if (_cssSettings.IncludeNormalFont)
            {
                var ns = _wb.Styles.GetNormalStyle();
                if (ns != null)
                {
                    WriteCssItem($"font-family:{ns.Style.Font.Name};", _settings.Minify);
                    WriteCssItem($"font-size:{ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", _settings.Minify);
                }
            }
            foreach (var item in _cssSettings.AdditionalCssElements)
            {
                WriteCssItem($"{item.Key}:{item.Value};", _settings.Minify);
            }
            WriteClassEnd(_settings.Minify);

            //Class for hidden rows.
            WriteClass($".{_settings.StyleClassPrefix}hidden {{", _settings.Minify);
            WriteCssItem($"display:none;", _settings.Minify);
            WriteClassEnd(_settings.Minify);

            WriteClass($".{_settings.StyleClassPrefix}al {{", _settings.Minify);
            WriteCssItem($"text-align:left;", _settings.Minify);
            WriteClassEnd(_settings.Minify);

            WriteClass($".{_settings.StyleClassPrefix}ar {{", _settings.Minify);
            WriteCssItem($"text-align:right;", _settings.Minify);
            WriteClassEnd(_settings.Minify);

            var worksheets=_ranges.Select(x=>x.Worksheet).Distinct().ToList();
            foreach (var ws in worksheets)
            {
                var clsName = HtmlExportTableUtil.GetWorksheetClassName(_settings.StyleClassPrefix, "dcw", ws, worksheets.Count > 1);
                WriteClass($".{clsName} {{", _settings.Minify);
                WriteCssItem($"width:{ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), ws.Workbook.MaxFontWidth)}px;", _settings.Minify);
                WriteClassEnd(_settings.Minify);

                clsName = HtmlExportTableUtil.GetWorksheetClassName(_settings.StyleClassPrefix, "drh", ws, worksheets.Count > 1);
                WriteClass($".{clsName} {{", _settings.Minify);
                WriteCssItem($"height:{(int)(ws.DefaultRowHeight / 0.75)}px;", _settings.Minify);
                WriteClassEnd(_settings.Minify);
            }

            //Image alignment class
            if (_settings.Pictures.Include != ePictureInclude.Exclude && _settings.Pictures.CssExclude.Alignment == false)
            {
                WriteClass($"td.{_settings.StyleClassPrefix}image-cell {{", _settings.Minify);
                if (_settings.Pictures.AddMarginTop)
                {
                    WriteCssItem($"vertical-align:top;", _settings.Minify);
                }
                else
                {
                    WriteCssItem($"vertical-align:middle;", _settings.Minify);
                }
                if (_settings.Pictures.AddMarginTop)
                {
                    WriteCssItem($"text-align:left;", _settings.Minify);
                }
                else
                {
                    WriteCssItem($"text-align:center;", _settings.Minify);
                }
                WriteClassEnd(_settings.Minify);
            }
        }
        internal void AddPictureToCss(HtmlImage p)
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
                WriteClass($"img.{_settings.StyleClassPrefix}image-{imageFileName}{{", _settings.Minify);
                WriteCssItem($"content:url('data:{GetContentType(type.Value)};base64,{encodedImage}');", _settings.Minify);
                if(_settings.Pictures.Position!=ePicturePosition.DontSet)
                {
                    WriteCssItem($"position:{_settings.Pictures.Position.ToString().ToLower()};", _settings.Minify);
                }

                if (p.FromColumnOff != 0 && _settings.Pictures.AddMarginLeft)
                {
                    var leftOffset = p.FromColumnOff / ExcelPicture.EMU_PER_PIXEL;
                    WriteCssItem($"margin-left:{leftOffset}px;", _settings.Minify);
                }

                if (p.FromRowOff != 0 && _settings.Pictures.AddMarginTop)
                {
                    var topOffset = p.FromRowOff / ExcelPicture.EMU_PER_PIXEL;
                    WriteCssItem($"margin-top:{topOffset}px;", _settings.Minify);
                }


                WriteClassEnd(_settings.Minify);
                _images.Add(pc.ImageHash);
            }
            AddPicturePropertiesToCss(p);
        }

        private void AddPicturePropertiesToCss(HtmlImage image)
        {
            string imageName = HtmlExportTableUtil.GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);
            var width = image.Picture.GetPixelWidth();
            var height = image.Picture.GetPixelHeight();

            WriteClass($"img.{_settings.StyleClassPrefix}image-prop-{imageName}{{", _settings.Minify);
            
            if (_settings.Pictures.KeepOriginalSize == false)
            {
                if (width != image.Picture.Image.Bounds.Width)
                {
                    WriteCssItem($"max-width:{width:F0}px;", _settings.Minify);
                }
                if (height != image.Picture.Image.Bounds.Height)
                {
                    WriteCssItem($"max-height:{height:F0}px;", _settings.Minify);
                }
            }

            if(image.Picture.Border.LineStyle!=null && _settings.Pictures.CssExclude.Border == false)
            {
                var border = GetDrawingBorder(image.Picture);
                WriteCssItem($"border:{border};", _settings.Minify);
            }
            WriteClassEnd(_settings.Minify);
        }

        private string GetDrawingBorder(ExcelPicture picture)
        {
            Color color = picture.Border.Fill.Color;
            if (color.IsEmpty) return "";
            string lineStyle=$"{picture.Border.Width}px";
            
            switch (picture.Border.LineStyle.Value)
            {
                case eLineStyle.Solid:
                    lineStyle += " solid";
                    break;
                case eLineStyle.Dash:
                case eLineStyle.LongDashDot:
                case eLineStyle.LongDashDotDot:
                case eLineStyle.SystemDash:
                case eLineStyle.SystemDashDot:
                case eLineStyle.SystemDashDotDot:
                    lineStyle += $" dashed";
                    break;
                case eLineStyle.Dot:
                    lineStyle += $" dot";
                    break;
            }

            lineStyle += " #" + color.ToArgb().ToString("x8").Substring(2);
            return lineStyle;
        }

        private object GetContentType(ePictureType type)
        {
            switch(type)
            {
                case ePictureType.Ico:
                    return "image/vnd.microsoft.icon";
                case ePictureType.Jpg:
                    return "image/jpeg";
                case ePictureType.Svg:
                    return "image/svg+xml";
                case ePictureType.Tif:
                    return "image/tiff";
                default:
                    return $"image/{type}";
            }
        }
        internal void AddToCss(ExcelStyles styles, int styleId, string styleClassPrefix, string cellStyleClassName)
        {
            var xfs = styles.CellXfs[styleId];
            if (HasStyle(xfs))
            {
                if (IsAddedToCache(xfs, out int id)==false || _addedToCss.Contains(id) == false)
                {
                    _addedToCss.Add(id);
                    WriteClass($".{styleClassPrefix}{cellStyleClassName}{id}{{", _settings.Minify);
                    if (xfs.FillId > 0)
                    {
                        WriteFillStyles(xfs.Fill);
                    }
                    if (xfs.FontId > 0)
                    {
                        var ns = styles.GetNormalStyle();
                        WriteFontStyles(xfs.Font, ns.Style.Font);
                    }
                    if (xfs.BorderId > 0)
                    {
                        WriteBorderStyles(xfs.Border.Top, xfs.Border.Bottom, xfs.Border.Left, xfs.Border.Right);
                    }
                    WriteStyles(xfs);
                    WriteClassEnd(_settings.Minify);
                }
            }
        }

        internal void AddToCss(ExcelStyles styles, int styleId, int bottomStyleId, int rightStyleId, string styleClassPrefix, string cellStyleClassName)
        {
            var xfs = styles.CellXfs[styleId];
            var bXfs = styles.CellXfs[bottomStyleId];
            var rXfs = styles.CellXfs[rightStyleId];
            if (HasStyle(xfs) || bXfs.BorderId > 0 || rXfs.BorderId > 0)
            {
                if (IsAddedToCache(xfs, out int id, bottomStyleId, rightStyleId) == false || _addedToCss.Contains(id) == false)
                {
                    _addedToCss.Add(id);
                    WriteClass($".{styleClassPrefix}{cellStyleClassName}{id}{{", _settings.Minify);
                    if (xfs.FillId > 0)
                    {
                        WriteFillStyles(xfs.Fill);
                    }
                    if (xfs.FontId > 0)
                    {
                        var ns = styles.GetNormalStyle();
                        WriteFontStyles(xfs.Font, ns.Style.Font);
                    }
                    if (xfs.BorderId > 0 || bXfs.BorderId > 0 || rXfs.BorderId > 0)
                    {
                        WriteBorderStyles(xfs.Border.Top, bXfs.Border.Bottom, xfs.Border.Left, rXfs.Border.Right);
                    }
                    WriteStyles(xfs);
                    WriteClassEnd(_settings.Minify);
                }
            }
        }

        private bool IsAddedToCache(ExcelXfs xfs, out int id, int bottomStyleId = -1, int rightStyleId = -1)
        {
            var key = GetStyleKey(xfs);
            if (bottomStyleId > -1) key += bottomStyleId + "|" + rightStyleId;
            if (_styleCache.ContainsKey(key))
            {
                id = _styleCache[key];
                return true;
            }
            else
            {
                id = _styleCache.Count+1;
                _styleCache.Add(key, id);
                return false;
            }
        }

        private void WriteStyles(ExcelXfs xfs)
        {
            if (_cssExclude.WrapText == false)
            {
                if (xfs.WrapText)
                {
                    WriteCssItem("white-space: break-spaces;", _settings.Minify);
                }
                else
                {
                    WriteCssItem("white-space: nowrap;", _settings.Minify);
                }
            }

            if (xfs.HorizontalAlignment != ExcelHorizontalAlignment.General && _cssExclude.HorizontalAlignment == false)
            {
                var hAlign = GetHorizontalAlignment(xfs);
                WriteCssItem($"text-align:{hAlign};", _settings.Minify);
            }

            if (xfs.VerticalAlignment != ExcelVerticalAlignment.Bottom && _cssExclude.VerticalAlignment == false)
            {
                var vAlign = GetVerticalAlignment(xfs);
                WriteCssItem($"vertical-align:{vAlign};", _settings.Minify);
            }
            if(xfs.TextRotation!=0 && _cssExclude.TextRotation==false)
            {
                if(xfs.TextRotation==255)
                {
                    WriteCssItem($"writing-mode:vertical-lr;;", _settings.Minify);
                    WriteCssItem($"text-orientation:upright;", _settings.Minify);                    
                }
                else
                {                    
                    if(xfs.TextRotation>90)
                    {
                        WriteCssItem($"transform:rotate({xfs.TextRotation-90}deg);", _settings.Minify);
                    }
                    else
                    {
                        WriteCssItem($"transform:rotate({360-xfs.TextRotation}deg);", _settings.Minify);
                    }
                }
            }

            if(xfs.Indent > 0 && _cssExclude.Indent == false)
            {
                WriteCssItem($"padding-left:{xfs.Indent * _cssSettings.IndentValue}{_cssSettings.IndentUnit};", _settings.Minify);
            }
        }

        private void WriteBorderStyles(ExcelBorderItemXml top, ExcelBorderItemXml bottom, ExcelBorderItemXml left, ExcelBorderItemXml right)
        {
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Top)) WriteBorderItem(top, "top");
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Bottom)) WriteBorderItem(bottom, "bottom");
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Left)) WriteBorderItem(left, "left");
            if (EnumUtil.HasNotFlag(_borderExclude, eBorderExclude.Right)) WriteBorderItem(right, "right");
            //TODO add Diagonal
            //WriteBorderItem(b.DiagonalDown, "right");
            //WriteBorderItem(b.DiagonalUp, "right");
        }

        private void WriteBorderItem(ExcelBorderItemXml bi, string suffix)
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

                WriteCssItem(sb.ToString(), _settings.Minify);
            }
        }

        private void WriteFontStyles(ExcelFontXml f, ExcelFont nf)
        {
            if(string.IsNullOrEmpty(f.Name)==false && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Name) && f.Name.Equals(nf.Name) == false)
            {
                WriteCssItem($"font-family:{f.Name};", _settings.Minify);
            }
            if(f.Size>0 && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Size) && f.Size!=nf.Size)
            {
                WriteCssItem($"font-size:{f.Size.ToString("g", CultureInfo.InvariantCulture)}pt;", _settings.Minify);
            }
            if (f.Color!=null && f.Color.Exists && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Color) && AreColorEqual(f.Color, nf.Color)==false)
            {
                WriteCssItem($"color:{GetColor(f.Color)};", _settings.Minify);
            }
            if (f.Bold && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Bold) && nf.Bold!=f.Bold)
            {
                WriteCssItem("font-weight:bolder;", _settings.Minify);
            }
            if (f.Italic && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Italic) && nf.Italic != f.Italic)
            {
                WriteCssItem("font-style:italic;", _settings.Minify);
            }
            if (f.Strike && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Strike) && nf.Strike != f.Strike)
            {
                WriteCssItem("text-decoration:line-through solid;", _settings.Minify);
            }
            if (f.UnderLineType != ExcelUnderLineType.None && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Underline) && f.UnderLineType!=nf.UnderLineType)
            {
                switch (f.UnderLineType)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        WriteCssItem("text-decoration:underline double;", _settings.Minify);
                        break;
                    default:
                        WriteCssItem("text-decoration:underline solid;", _settings.Minify);
                        break;
                }
            }            
        }

        private bool AreColorEqual(ExcelColorXml c1, ExcelColor c2)
        {
            if (c1.Tint != c2.Tint) return false;
            if(c1.Indexed>=0)
            {
                return c1.Indexed == c2.Indexed;
            }
            else if(string.IsNullOrEmpty(c1.Rgb)==false)
            {
                return c1.Rgb == c2.Rgb;
            }
            else if(c1.Theme!=null)
            {
                return c1.Theme == c2.Theme;
            }
            else
            {
                return c1.Auto == c2.Auto;
            }
        }

        private void WriteFillStyles(ExcelFillXml f)
        {
            if (_cssExclude.Fill) return;
            if (f is ExcelGradientFillXml gf && gf.Type!=ExcelFillGradientType.None)
            {
                WriteGradient(gf);
            }
            else
            {
                if (f.PatternType == ExcelFillStyle.Solid)
                {
                    WriteCssItem($"background-color:{GetColor(f.BackgroundColor)};", _settings.Minify);
                }
                else
                {
                    WriteCssItem($"{PatternFills.GetPatternSvg(f.PatternType, GetColor(f.BackgroundColor), GetColor(f.PatternColor))}", _settings.Minify);
                }
            }
        }

        private void WriteGradient(ExcelGradientFillXml gradient)
        {
            if (gradient.Type == ExcelFillGradientType.Linear)
            {
                _writer.Write($"background: linear-gradient({(gradient.Degree + 90) % 360}deg");
            }
            else
            {
                _writer.Write($"background:radial-gradient(ellipse {gradient.Right  * 100}% {gradient.Bottom  * 100}%");
            }

            _writer.Write($",{GetColor(gradient.GradientColor1)} 0%");
            _writer.Write($",{GetColor(gradient.GradientColor2)} 100%");

            _writer.Write(");");
        }
        private string GetColor(ExcelColorXml c)
        {
            Color ret;
            if (!string.IsNullOrEmpty(c.Rgb))
            {
                if (int.TryParse(c.Rgb, NumberStyles.HexNumber, null, out int hex))
                {
                    ret = Color.FromArgb(hex);
                }
                else
                {
                    ret = Color.Empty;
                }
            }
            else if (c.Theme.HasValue)
            {
                ret = ColorConverter.GetThemeColor(_theme, c.Theme.Value);
            }
            else if (c.Indexed >= 0)
            {
                ret = ExcelColor.GetIndexedColor(c.Indexed);
            }
            else
            {
                //Automatic, set to black.
                ret = Color.Black;
            }
            if (c.Tint != 0)
            {
                ret = ColorConverter.ApplyTint(ret, Convert.ToDouble(c.Tint));
            }
            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }
        public void FlushStream()
        {
            _writer.Flush();
        }
    }
}
