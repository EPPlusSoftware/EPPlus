using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Text;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Export.HtmlExport.Translators;

namespace OfficeOpenXml.Export.HtmlExport.Parsers
{
   internal partial class CssRangeTranslator
    {
        HtmlExportSettings _settings;
        CssExportSettings _cssSettings;
        CssExclude _cssExclude;

        ExcelWorkbook _wb;
        List<ExcelRangeBase> _ranges;
        ExcelTheme _theme;
        eBorderExclude _borderExclude;
        eFontExclude _fontExclude;

        CssRuleCollection _ruleCollection;

        internal CssRuleCollection RuleCollection => _ruleCollection;

        TranslatorContext _context;


        internal CssRangeTranslator(List<ExcelRangeBase> ranges, HtmlRangeExportSettings settings)
        {
            _settings = settings;
            _cssSettings = settings.Css;
            _cssExclude = settings.Css.CssExclude;
            Init(ranges);
            _ruleCollection = new CssRuleCollection();

            _context = new TranslatorContext();
            _context.FontExclude = _cssExclude.Font;
            _context.BorderExclude = _cssExclude.Border;
            _context.FillExclude = _cssExclude.Fill;
            _context.Theme = _theme;
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

        //internal void AddRuleToList(List<CssRule> list, string ruleName, string declarationName, params string[] declarationValues)
        //{
        //    var toBeAdded = new CssRule(ruleName)
        //    {
        //        Declarations =
        //        {
        //            new Declaration(declarationName, declarationValues),
        //        }
        //    };

        //    list.Add(toBeAdded);
        //}

        //internal void WriteRulesList(List<CssRule> rules) 
        //{
        //    for(int i = 0; i < rules.Count; i++) 
        //    {
        //        _cssTrueWriter.WriteRule(rules[i], _settings.Minify);
        //    }

        //    rules.Clear();
        //}

        internal void RenderAdditionalAndFontCss(string tableClass)
        {
            if (_cssSettings.IncludeSharedClasses == false) return;

            var tableRule = new CssRule($"table.{tableClass}");

            if (_cssSettings.IncludeNormalFont)
            {
                var ns = _wb.Styles.GetNormalStyle();
                if (ns != null)
                {
                    tableRule.Declarations.Add(new Declaration($"font-family", ns.Style.Font.Name));
                    tableRule.Declarations.Add(new Declaration($"font-size", $"{ns.Style.Font.Size.ToString("g", CultureInfo.InvariantCulture)}pt"));
                }
            }
            foreach (var item in _cssSettings.AdditionalCssElements)
            {
                //tableRule.Declarations.Add(new Declaration(item.Key,item.Value));
                tableRule.AddDeclaration(item.Key, item.Value);
            }

            // _cssTrueWriter.WriteRule(tableRule, _settings.Minify);

            List<CssRule> rulesToBeWritten = new List<CssRule>{ tableRule };

            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}hidden ", "display", "none");
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}al ", "text-align", "left");
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}ar ", "text-align", "right");

            //WriteRulesList(rulesToBeWritten);

            ////var hiddenClass = new CssRule($".{_settings.StyleClassPrefix}hidden")
            ////{
            ////    Declarations =
            ////    {
            ////        new Declaration("display", "none"),
            ////    }
            ////};

            ////var alignLeft = new CssRule($".{_settings.StyleClassPrefix}al")
            ////{
            ////    Declarations =
            ////    {
            ////        new Declaration("text-align", "left"),
            ////    }
            ////};

            ////var alignRight = new CssRule($".{_settings.StyleClassPrefix}ar")
            ////{
            ////    Declarations =
            ////    {
            ////        new Declaration("text-align", "left"),
            ////    }
            ////};

            ////List<CssRule> rulesToBeWritten = new List<CssRule>
            ////{
            ////    tableRule, hiddenClass, alignLeft, alignRight
            ////};

            //_cssTrueWriter.WriteRule(hiddenClass, _settings.Minify);
            //_cssTrueWriter.WriteRule(alignLeft, _settings.Minify);
            //_cssTrueWriter.WriteRule(alignRight, _settings.Minify);
        }

        internal void AddToCss(ExcelXfs xfs, ExcelNamedStyleXml ns, int id)
        {
            var styleClass = new CssRule($".{_settings.StyleClassPrefix}{_settings.CellStyleClassName}{id}");

            var translators = new List<TranslatorBase>();

            if (xfs.FillId > 0)
            {
                translators.Add(new CssFillTranslator(xfs.Fill));
            }
            if (xfs.FontId > 0)
            {
                translators.Add(new CssFontTranslator(xfs.Font, ns.Style.Font));
            }

            foreach(var translator in translators) 
            {
                _context.SetTranslator(translator);
                _context.AddDeclarations(styleClass);
            }

            AddGenericStyleDeclarations(xfs, styleClass);
            _ruleCollection.CssRules.Add(styleClass);
        }



        internal void AddToCss(List<ExcelXfs> xfsList, ExcelNamedStyleXml ns, int id)
        {
            var xfs = xfsList[0];
            var bXfs = xfsList[1];
            var rXfs = xfsList[2];

            if (/*HasStyle(xfs) ||*/ bXfs.BorderId > 0 || rXfs.BorderId > 0)
            {
                var styleClass = new CssRule($".{_settings.StyleClassPrefix}{_settings.CellStyleClassName}{id}");
                if (xfs.FillId > 0)
                {
                    AddFillDeclarations(xfs.Fill, styleClass);
                }

                //if (xfs.FontId > 0)
                //{
                //    WriteFontStyles(xfs.Font, ns.Style.Font);
                //}
                //if (xfs.BorderId > 0 || bXfs.BorderId > 0 || rXfs.BorderId > 0)
                //{
                //    WriteBorderStyles(xfs.Border.Top, bXfs.Border.Bottom, xfs.Border.Left, rXfs.Border.Right);
                //}
                //WriteStyles(xfs);
                //WriteClassEnd(_settings.Minify);

                 AddGenericStyleDeclarations(xfs, styleClass);
                _ruleCollection.CssRules.Add(styleClass);

                //_cssTrueWriter.WriteRule(styleClass, _settings.Minify);
            }
        }

        private void AddFillDeclarations(ExcelFillXml f, CssRule rule)
        {
            if (_cssExclude.Fill) return;
            if (f is ExcelGradientFillXml gf && gf.Type != ExcelFillGradientType.None)
            {
                WriteGradient(gf, rule);
            }
            else
            {
                if (f.PatternType == ExcelFillStyle.Solid)
                {
                    rule.AddDeclaration("background-color", GetColor(f.BackgroundColor));
                    //WriteCssItem($"background-color:{GetColor(f.BackgroundColor)};", _settings.Minify);
                }
                else
                {
                    var svg = PatternFills.GetPatternSvgConvertedOnly(f.PatternType, GetColor(f.BackgroundColor), GetColor(f.PatternColor));
                    rule.AddDeclaration("background-repeat", "repeat");
                    //arguably some of the values should be its own declaration...Should still work though.
                    rule.AddDeclaration("background", $"url(data:image/svg+xml;base64,{svg})");

                    //WriteCssItem($"{PatternFills.GetPatternSvg(f.PatternType, GetColor(f.BackgroundColor), GetColor(f.PatternColor))}", _settings.Minify);
                }
            }
        }

        private void WriteGradient(ExcelGradientFillXml gradient, CssRule rule)
        {
            if (gradient.Type == ExcelFillGradientType.Linear)
            {
                rule.AddDeclaration("background", $"linear-gradient({(gradient.Degree + 90) % 360}deg");
                //_writer.Write($"background: linear-gradient({(gradient.Degree + 90) % 360}deg");
            }
            else
            {
                rule.AddDeclaration("background", $"radial-gradient(ellipse {gradient.Right * 100}% {gradient.Bottom * 100}%"); 
                //_writer.Write($"background:radial-gradient(ellipse {gradient.Right * 100}% {gradient.Bottom * 100}%");
            }

            rule.Declarations.LastOrDefault().AddValues
                (
                $",{GetColor(gradient.GradientColor1)} 0%", 
                $",{GetColor(gradient.GradientColor2)} 100%)"
                );

            //_writer.Write($",{GetColor(gradient.GradientColor1)} 0%");
            //_writer.Write($",{GetColor(gradient.GradientColor2)} 100%");

            //_writer.Write(");");
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
                ret = Utils.ColorConverter.GetThemeColor(_theme, c.Theme.Value);
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
                ret = Utils.ColorConverter.ApplyTint(ret, Convert.ToDouble(c.Tint));
            }
            return "#" + ret.ToArgb().ToString("x8").Substring(2);
        }

        private void AddGenericStyleDeclarations(ExcelXfs xfs, CssRule rule)
        {
            if (_cssExclude.WrapText == false)
            {
                rule.AddDeclaration("white-space", xfs.WrapText ? "break-spaces" : "nowrap");

                //if (xfs.WrapText)
                //{
                //    rule.AddDeclaration("white-space", "break-spaces");
                //    //WriteCssItem("white-space: break-spaces;", _settings.Minify);
                //}
                //else
                //{
                //    rule.AddDeclaration("white-space", "nowrap");
                //}
            }

            if (xfs.HorizontalAlignment != ExcelHorizontalAlignment.General && _cssExclude.HorizontalAlignment == false)
            {
                //var hAlign = GetHorizontalAlignment(xfs);
                rule.AddDeclaration("text-align", GetHorizontalAlignment(xfs));
                //WriteCssItem($"text-align:{hAlign};", _settings.Minify);
            }

            if (xfs.VerticalAlignment != ExcelVerticalAlignment.Bottom && _cssExclude.VerticalAlignment == false)
            {
                rule.AddDeclaration("vertical-align", GetVerticalAlignment(xfs));

                //var vAlign = GetVerticalAlignment(xfs);
                //WriteCssItem($"vertical-align:{vAlign};", _settings.Minify);
            }
            if (xfs.TextRotation != 0 && _cssExclude.TextRotation == false)
            {
                if (xfs.TextRotation == 255)
                {
                    rule.AddDeclaration("writing-mode", "vertical-lr");
                    rule.AddDeclaration("text-orientation", "upright");

                    //WriteCssItem($"writing-mode:vertical-lr;;", _settings.Minify);
                    //WriteCssItem($"text-orientation:upright;", _settings.Minify);
                }
                else
                {
                    var rotationvalue = xfs.TextRotation > 90 ? xfs.TextRotation - 90 : 360 - xfs.TextRotation;
                    rule.AddDeclaration("transform", $"rotate({rotationvalue}deg)");

                    //if (xfs.TextRotation > 90)
                    //{
                    //    rule.AddDeclaration("transform", $"rotate({xfs.TextRotation - 90}deg)");
                    //    //WriteCssItem($"transform:rotate({xfs.TextRotation - 90}deg);", _settings.Minify);
                    //}
                    //else
                    //{
                    //    WriteCssItem($"transform:rotate({360 - xfs.TextRotation}deg);", _settings.Minify);
                    //}
                }
            }

            if (xfs.Indent > 0 && _cssExclude.Indent == false)
            {
                rule.AddDeclaration("transform", $"padding-left:{xfs.Indent * _cssSettings.IndentValue}{_cssSettings.IndentUnit}");
            }
        }

        private void AddFontDeclarations(ExcelFontXml f, ExcelFont nf, CssRule rule)
        {
            if (string.IsNullOrEmpty(f.Name) == false && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Name) && f.Name.Equals(nf.Name) == false)
            {
                rule.AddDeclaration("font-family", f.Name);
            }
            if (f.Size > 0 && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Size) && f.Size != nf.Size)
            {
                rule.AddDeclaration("font-size", $"{f.Size.ToString("g", CultureInfo.InvariantCulture)}pt");
            }
            if (f.Color != null && f.Color.Exists && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Color) && AreColorEqual(f.Color, nf.Color) == false)
            {
                rule.AddDeclaration("color", GetColor(f.Color));
            }
            if (f.Bold && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Bold) && nf.Bold != f.Bold)
            {
                rule.AddDeclaration("font-weight", "bolder");
            }
            if (f.Italic && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Italic) && nf.Italic != f.Italic)
            {
                rule.AddDeclaration("font-style", "italic");
            }
            if (f.Strike && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Strike) && nf.Strike != f.Strike)
            {
                rule.AddDeclaration("text-decoration", "line-through", "solid");
            }
            if (f.UnderLineType != ExcelUnderLineType.None && EnumUtil.HasNotFlag(_fontExclude, eFontExclude.Underline) && f.UnderLineType != nf.UnderLineType)
            {
                switch (f.UnderLineType)
                {
                    case ExcelUnderLineType.Double:
                    case ExcelUnderLineType.DoubleAccounting:
                        rule.AddDeclaration("text-decoration", "underline", "double");
                        break;
                    default:
                        rule.AddDeclaration("text-decoration", "underline", "solid");
                        break;
                }
            }
        }

        private bool AreColorEqual(ExcelColorXml c1, ExcelColor c2)
        {
            if (c1.Tint != c2.Tint) return false;
            if (c1.Indexed >= 0)
            {
                return c1.Indexed == c2.Indexed;
            }
            else if (string.IsNullOrEmpty(c1.Rgb) == false)
            {
                return c1.Rgb == c2.Rgb;
            }
            else if (c1.Theme != null)
            {
                return c1.Theme == c2.Theme;
            }
            else
            {
                return c1.Auto == c2.Auto;
            }
        }

        protected static string GetVerticalAlignment(ExcelXfs xfs)
        {
            switch (xfs.VerticalAlignment)
            {
                case ExcelVerticalAlignment.Top:
                    return "top";
                case ExcelVerticalAlignment.Center:
                    return "middle";
                case ExcelVerticalAlignment.Bottom:
                    return "bottom";
            }

            return "";
        }

        protected static string GetHorizontalAlignment(ExcelXfs xfs)
        {
            switch (xfs.HorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Right:
                    return "right";
                case ExcelHorizontalAlignment.Center:
                case ExcelHorizontalAlignment.CenterContinuous:
                    return "center";
                case ExcelHorizontalAlignment.Left:
                    return "left";
            }

            return "";
        }


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
