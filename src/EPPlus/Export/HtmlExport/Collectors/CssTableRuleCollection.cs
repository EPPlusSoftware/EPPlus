using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Contexts;
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

            _context = new TranslatorContext(_settings.Css.Exclude.TableStyle);
            _context.Theme = _theme;
        }

        internal void AddHyperlink(string name, ExcelTableStyleElement element)
        {
            var styleClass = new CssRule($"table.{name} a");

            var ft = new CssFontTranslator(new FontDxf(element.Style.Font), null);

            styleClass.AddDeclarationList(ft.GenerateDeclarationList(_context));

            _ruleCollection.AddRule(styleClass);
        }

        internal void AddAlignment(string name, List<string> dataTypes)
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

                bool rightDefault;
                //We never want to default horizontal right if horizontal is excluded.
                if (_context.Exclude.HorizontalAlignment == false)
                {
                    rightDefault = false;
                }
                else
                {
                    rightDefault = c < dataTypes.Count && 
                        (dataTypes[c] == HtmlDataTypes.Number || dataTypes[c] == HtmlDataTypes.DateTime);
                }

                if (rightDefault || styleId > 0)
                {
                    var styleClass = new CssRule($"table.{name} td:nth-child({col})");

                    if (styleId > 0)
                    {
                        var xfs = new StyleXml(_table.WorkSheet.Workbook.Styles.CellXfs[styleId]);
                        var translator = new CssTableTextFormatTranslator(xfs, rightDefault);

                        styleClass.AddDeclarationList(translator.GenerateDeclarationList(_context));
                    }
                    else
                    {
                        styleClass.AddDeclaration("text-align", "right");
                    }

                    //TODO: If we exclude both horizontal and vertical we can get a class with empty declaration list here...
                    _ruleCollection.AddRule(styleClass);
                }
            }
        }
    }
}
