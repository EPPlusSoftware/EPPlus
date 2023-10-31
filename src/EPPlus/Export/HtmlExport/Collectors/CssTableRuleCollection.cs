using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static OfficeOpenXml.Export.HtmlExport.ColumnDataTypeManager;

namespace OfficeOpenXml.Export.HtmlExport.Collectors
{
    internal class CssTableRuleCollection
    {
        protected HtmlTableExportSettings _settings;
        ExcelTable _table;
        ExcelTheme _theme;

        CssRuleCollection _ruleCollection;
        TranslatorContext _context;

        internal CssTableRuleCollection(ExcelTable table, HtmlTableExportSettings settings)
        {
            _table = table;
            _settings = settings;
            if (table.WorkSheet.Workbook.ThemeManager.CurrentTheme == null)
            {
                table.WorkSheet.Workbook.ThemeManager.CreateDefaultTheme();
            }
            _theme = table.WorkSheet.Workbook.ThemeManager.CurrentTheme;

            _context.Theme = _theme;
        }

        internal void AddHyperlink(string name, ExcelTableStyleElement element)
        {
            var styleClass = new CssRule($"table.{name} a");

            var ft = new CssFontTranslator(new FontDxf(element.Style.Font), null);

            styleClass.AddDeclarationList(ft.GenerateDeclarationList(_context));

            _ruleCollection.AddRule(styleClass);
        }

        internal void AddAlignmentToCss(string name, List<string> dataTypes)
        {
            if (_settings.HorizontalAlignmentWhenGeneral == eHtmlGeneralAlignmentHandling.DontSet)
            {
                return;
            }
            var row = _table.ShowHeader ? _table.Address._fromRow + 1 : _table.Address._fromRow;
            for (int c = 0; c < _table.Columns.Count; c++)
            {
                var col = _table.Address._fromCol + c;
                var styleId = _table.WorkSheet.GetStyleInner(row, col);
                string hAlign = "";
                string vAlign = "";
                if (styleId > 0)
                {
                    var xfs = _table.WorkSheet.Workbook.Styles.CellXfs[styleId];
                    if (xfs.ApplyAlignment ?? false)
                    {
                        hAlign = GetHorizontalAlignment(xfs);
                        vAlign = GetVerticalAlignment(xfs);
                    }
                }

                if (string.IsNullOrEmpty(hAlign) && c < dataTypes.Count && (dataTypes[c] == HtmlDataTypes.Number || dataTypes[c] == HtmlDataTypes.DateTime))
                {
                    hAlign = "right";
                }

                if (!(string.IsNullOrEmpty(hAlign) && string.IsNullOrEmpty(vAlign)))
                {
                    WriteClass($"table.{name} td:nth-child({col}){{", _settings.Minify);
                    if (string.IsNullOrEmpty(hAlign) == false && _settings.Css.Exclude.TableStyle.HorizontalAlignment == false)
                    {
                        WriteCssItem($"text-align:{hAlign};", _settings.Minify);
                    }
                    if (string.IsNullOrEmpty(vAlign) == false && _settings.Css.Exclude.TableStyle.VerticalAlignment == false)
                    {
                        WriteCssItem($"vertical-align:{vAlign};", _settings.Minify);
                    }
                    WriteClassEnd(_settings.Minify);
                }
            }
        }
}
