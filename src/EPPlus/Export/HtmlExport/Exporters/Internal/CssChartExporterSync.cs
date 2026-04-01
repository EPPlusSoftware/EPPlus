using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Export.HtmlExport.Determinator;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace OfficeOpenXml.Export.HtmlExport.Exporters.Internal
{
    internal class CssChartExporterSync
    {
        CssExportSettings _settings;

        ExcelChart _chart;

        TranslatorContext _context;

        public CssChartExporterSync(ExcelChart chart) : this(new CssChartExportSettings(), chart)
        {

        }

        public CssChartExporterSync(CssChartExportSettings settings, ExcelChart chart)
        {
            _settings = settings;
            Require.Argument(chart).IsNotNull("chart");
            _chart = chart;
        }
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetCssString()
        {
            using (var ms = EPPlusMemoryManager.GetStream())
            {
                RenderCss(ms);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        /// <summary>
        /// Exports the css part of the html export.
        /// </summary>
        /// <param name="stream">The stream to write the css to.</param>
        /// <exception cref="IOException"></exception>
        public void RenderCss(Stream stream)
        {
            var trueWriter = new CssWriter(stream);

            var cssCollection = new CssChartRuleCollection(_chart, _settings);

            trueWriter.WriteAndClearFlush(cssCollection, false);
        }

        //    /// <summary>
        //    /// Exports the css part of an <see cref="ExcelTable"/> to a html string
        //    /// </summary>
        //    /// <returns>A html table</returns>
        //    public void RenderCss(Stream stream)
        //    {
        //        var cssWriter = GetTableCssWriter(stream, _table, _tableSettings);
        //        if (cssWriter == null) { return; }

        //        var cssRules = CreateRuleCollection(_tableSettings);
        //        cssWriter.WriteAndClearFlush(cssRules, Settings.Minify);
        //    }

        //    protected CssRuleCollection CreateRuleCollection(CssExportSettings settings)
        //    {
        //        var cssTranslator = new CssChartRuleCollection(_chart, settings);

        //        _context = new TranslatorContext(settings);


        //        //AddCssRulesToCollection(cssTranslator, settings);

        //        //return cssTranslator.RuleCollection;
        //    }

        }
    }
