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

            _context = new TranslatorContext(settings.Css.CssExclude);
            _context.Theme = _theme;
            _context.IndentValue = _cssSettings.IndentValue;
            _context.IndentUnit = _cssSettings.IndentUnit;
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
                tableRule.AddDeclaration(item.Key, item.Value);
            }

            List<CssRule> rulesToBeWritten = new List<CssRule> { tableRule };

            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}hidden ", "display", "none");
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}al ", "text-align", "left");
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}ar ", "text-align", "right");
        }

        internal void AddToCollection(ExcelXfs xfs, ExcelNamedStyleXml ns, int id)
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
            if (xfs.BorderId > 0)
            {
                translators.Add(new CssBorderTranslator(xfs.Border));
            }

            translators.Add(new CssTextFormatTranslator(xfs));

            foreach (var translator in translators)
            {
                _context.SetTranslator(translator);
                _context.AddDeclarations(styleClass);
            }

            _ruleCollection.CssRules.Add(styleClass);
        }

        internal void AddToCollection(List<ExcelXfs> xfsList, ExcelNamedStyleXml ns, int id)
        {
            var xfs = xfsList[0];
            var bXfs = xfsList[1];
            var rXfs = xfsList[2];

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
            if (xfs.BorderId > 0 || bXfs.BorderId > 0 || rXfs.BorderId > 0)
            {
                translators.Add(new CssBorderTranslator(xfs.Border.Top, bXfs.Border.Bottom, xfs.Border.Left, rXfs.Border.Right));
            }

            translators.Add(new CssTextFormatTranslator(xfs));

            foreach (var translator in translators)
            {
                _context.SetTranslator(translator);
                _context.AddDeclarations(styleClass);
            }

            _ruleCollection.CssRules.Add(styleClass);
        }
    }
}
