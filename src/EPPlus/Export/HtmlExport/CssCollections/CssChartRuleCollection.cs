using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Exporters.Internal;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.CssCollections
{
    internal class CssChartRuleCollection : CssRuleCollection
    {
        CssExportSettings _settings;
        ExcelChart _chart;
        ExcelTheme _theme;
        TranslatorContext _context;

        public CssChartRuleCollection(ExcelChart chart, CssExportSettings settings)
        {
            _chart = chart;
            _settings = settings;

            if (_chart.WorkSheet.Workbook.ThemeManager.CurrentTheme == null)
            {
                _chart.WorkSheet.Workbook.ThemeManager.CreateDefaultTheme();
            }
            _theme = _chart.WorkSheet.Workbook.ThemeManager.CurrentTheme;

            _context = new TranslatorContext(settings);
            _context.Theme = _theme;

            //RuleCollection = new CssRuleCollection();

            AddChartFillToCollection(chart, "epp-");
        }

        internal void AddChartFillToCollection(ExcelChart chart, string chartClassPreset)
        {
            var chartClass = $"{chartClassPreset}{HtmlExportTableUtil.GetClassName(chart.Name, $"chartstyle{chart.Id}")}";

            var s = chart.StyleManager.Style;

            if(s == null)
            {
                chart.StyleManager.ApplyStyles();
            }

            s = chart.StyleManager.Style;

            if (s != null)
            {

                var chartAreaRef = chart.StyleManager.Style.ChartArea.FillReference.Color;

                AddToCollection(chartClass, chart.Fill, chart.Border, "rect");


                //var fallbackElementBorder = chart.StyleManager.Style.ChartArea.BorderReference.Color;

                //AddToCollection(chartClass, chart.Border, "border");
            }
        }

        internal void AddToCollection(string name, ExcelDrawingBorder element, string htmlElement)
        {
            //if (element) return; //Dont add empty elements

            var s = element;

            var styleClass = new CssRule($"{htmlElement}.{name}", int.MaxValue);

            var translators = new List<TranslatorBase>();

            if (element != null && _context.Exclude.Fill == false)
            {
                //TODO: Ensure if gradients with more than 2 colors it is handled correctly.
                translators.Add(new CssFillTranslator(new FillDrawingBasic(element.Fill)));
            }
            //if (s.Font != null && _context.Exclude.Font != eFontExclude.All)
            //{
            //    translators.Add(new CssFontTranslator(new FontDxf(s.Font), null));
            //}
            //if (s.Border != null && _context.Exclude.Border != eBorderExclude.All)
            //{
            //    translators.Add(new CssBorderTranslator(new BorderDxf(s.Border)));
            //}

            foreach (var translator in translators)
            {
                _context.SetTranslator(translator);
                _context.AddDeclarations(styleClass);
            }

            AddRule(styleClass);
        }

        internal void AddToCollection(string name, ExcelDrawingFill element, ExcelDrawingBorder border, string htmlElement)
        {
            if (element.IsEmpty) return; //Dont add empty elements

            var s = element;

            var styleClass = new CssRule($"{htmlElement}.{name}", int.MaxValue);

            var translators = new List<TranslatorBase>();

            if (element != null && _context.Exclude.Fill == false)
            {
                var fillGeneric = new FillDrawing(element);
                var fillTranslator = new CssFillTranslator(fillGeneric, true);
                //TODO: Ensure if gradients with more than 2 colors it is handled correctly.
                translators.Add(fillTranslator);
            }
            //if (s.Font != null && _context.Exclude.Font != eFontExclude.All)
            //{
            //    translators.Add(new CssFontTranslator(new FontDxf(s.Font), null));
            //}
            if (border != null && _context.Exclude.Border != eBorderExclude.All)
            {
                
                translators.Add(new CssStrokeTranslator(new BorderDrawing(border)));
            }

            foreach (var translator in translators)
            {
                _context.SetTranslator(translator);
                _context.AddDeclarations(styleClass);
            }

            AddRule(styleClass);
        }

        internal void AddToCollection(string name, ExcelChartStyleItem element, string htmlElement)
        {
            if (element.HasValue() == false) return; //Dont add empty elements

            var s = element;

            var styleClass = new CssRule($"{htmlElement}.{name}", int.MaxValue);

            var translators = new List<TranslatorBase>();

            if (s.Fill != null && _context.Exclude.Fill == false)
            {
                translators.Add(new CssFillTranslator(new FillDrawing(s.Fill)));
            }
            //if (s.Font != null && _context.Exclude.Font != eFontExclude.All)
            //{
            //    translators.Add(new CssFontTranslator(new FontDxf(s.Font), null));
            //}
            //if (s.Border != null && _context.Exclude.Border != eBorderExclude.All)
            //{
            //    translators.Add(new CssBorderTranslator(new BorderDxf(s.Border)));
            //}

            foreach (var translator in translators)
            {
                _context.SetTranslator(translator);
                _context.AddDeclarations(styleClass);
            }

            AddRule(styleClass);
        }
    }
}
