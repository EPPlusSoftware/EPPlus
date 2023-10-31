using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Export.HtmlExport.Exporters;
using OfficeOpenXml.Drawing.Interfaces;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;

namespace OfficeOpenXml.Export.HtmlExport.Collectors
{
    internal partial class CssRangeRuleCollection
    {
        HtmlExportSettings _settings;
        CssExportSettings _cssSettings;

        ExcelWorkbook _wb;
        List<ExcelRangeBase> _ranges;
        ExcelTheme _theme;

        CssRuleCollection _ruleCollection;

        internal protected HashSet<string> _images = new HashSet<string>();

        internal CssRuleCollection RuleCollection => _ruleCollection;

        TranslatorContext _context;

        internal CssRangeRuleCollection(List<ExcelRangeBase> ranges, HtmlRangeExportSettings settings)
        {
            _settings = settings;
            _cssSettings = settings.Css;
            Init(ranges);
            _ruleCollection = new CssRuleCollection();

            _context = new TranslatorContext(settings);
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
        }

        internal void AddSharedClasses(string tableClass)
        {
            if (_cssSettings.IncludeSharedClasses == false) return;

            AddTableRule(tableClass);

            //Hidden class
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}hidden ", "display", "none");
            //Text-alignment classes
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}al ", "text-align", "left");
            _ruleCollection.AddRule($".{_settings.StyleClassPrefix}ar ", "text-align", "right");

            AddWorksheetDimensions();
            AddImageAlignment();
        }

        private void AddTableRule(string tableClass)
        {
            var tableRule = new CssRule($"table.{tableClass}");

            _context.SetTranslator(new CssTableTranslator(_wb.Styles.GetNormalStyle()));
            _context.AddDeclarations(tableRule);
            _ruleCollection.AddRule(tableRule);
        }

        private void AddImageAlignment()
        {
            if (_settings.Pictures.Include != ePictureInclude.Exclude && _settings.Pictures.CssExclude.Alignment == false)
            {
                var imgClass = new CssRule($"td.{_settings.StyleClassPrefix}image-cell ");
                imgClass.AddDeclaration("vertical-align", _settings.Pictures.AddMarginTop ? "top" : "middle");
                imgClass.AddDeclaration("text-align", _settings.Pictures.AddMarginLeft ? "left" : "center");

                _ruleCollection.AddRule(imgClass);
            }
        }

        private void AddWorksheetDimensions()
        {
            var worksheets = _ranges.Select(x => x.Worksheet).Distinct().ToList();
            foreach (var ws in worksheets)
            {
                var clsName = "." + HtmlExportTableUtil.GetWorksheetClassName(_settings.StyleClassPrefix, "dcw", ws, worksheets.Count > 1);
                CssRule widthRule = new CssRule(clsName);
                widthRule.AddDeclaration("width", $"{ExcelColumn.ColumnWidthToPixels(Convert.ToDecimal(ws.DefaultColWidth), ws.Workbook.MaxFontWidth)}px");

                clsName = "." + HtmlExportTableUtil.GetWorksheetClassName(_settings.StyleClassPrefix, "drh", ws, worksheets.Count > 1);
                CssRule heightRule = new CssRule(clsName);
                heightRule.AddDeclaration("height", $"{(int)(ws.DefaultRowHeight / 0.75)}px");

                _ruleCollection.AddRule(widthRule);
                _ruleCollection.AddRule(heightRule);
            }
        }

        internal void AddToCollection(List<IStyleExport> styleList, ExcelNamedStyleXml ns, int id)
        {
            var style = styleList[0];

            var styleClass = new CssRule($".{_settings.StyleClassPrefix}{_settings.CellStyleClassName}{id}");
            var translators = new List<TranslatorBase>();

            if (style.Fill != null)
            {
                translators.Add(new CssFillTranslator(style.Fill));
            }
            if (style.Font != null)
            {
                translators.Add(new CssFontTranslator(style.Font, ns.Style.Font));
            }

            if (styleList.Count > 1)
            {
                var bXfs = styleList[1];
                var rXfs = styleList[2];

                if (style.Border != null || bXfs.Border != null || rXfs.Border != null)
                {
                    translators.Add(new CssBorderTranslator(style.Border.Top, bXfs.Border.Bottom, style.Border.Left, rXfs.Border.Right));
                }
            }
            else if (style.Border != null)
            {
                translators.Add(new CssBorderTranslator(style.Border));
            }

            foreach (var translator in translators)
            {
                _context.SetTranslator(translator);
                _context.AddDeclarations(styleClass);
            }

            _ruleCollection.CssRules.Add(styleClass);
        }

        internal void AddPictureToCss(HtmlImage p)
        {
            var translator = new CssImageTranslator(p);

            if (translator.type == null) return;

            var pc = (IPictureContainer)p.Picture;

            if (_images.Contains(pc.ImageHash) == false)
            {
                string imageFileName = HtmlExportImageUtil.GetPictureName(p);
                var imgRule = new CssRule($"img.{_settings.StyleClassPrefix}image-{imageFileName}");

                _context.SetTranslator(translator);
                _context.AddDeclarations(imgRule);

                _ruleCollection.AddRule(imgRule);
                _images.Add(pc.ImageHash);
            }
            AddPicturePropertiesToCss(p);
        }

        private void AddPicturePropertiesToCss(HtmlImage image)
        {
            string imageName = HtmlExportTableUtil.GetClassName(image.Picture.Name, ((IPictureContainer)image.Picture).ImageHash);

            var imgProperties = new CssRule($"img.{_settings.StyleClassPrefix}image-prop-{imageName}");
            _context.SetTranslator(new CssImagePropertiesTranslator(image));
            _context.AddDeclarations(imgProperties);

            RuleCollection.AddRule(imgProperties);
        }
    }
}
