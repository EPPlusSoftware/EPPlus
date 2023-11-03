using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
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

        internal CssRuleCollection RuleCollection { get; private set; }
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

            RuleCollection = new CssRuleCollection();
        }

        internal void AddHyperlink(string name, ExcelTableStyleElement element)
        {
            if(_context.Exclude.Font != eFontExclude.All && 
                element.Style.HasValue && 
                element.Style.Font.HasValue)
            {
                var styleClass = new CssRule($"table.{name} a");

                var ft = new CssFontTranslator(new FontDxf(element.Style.Font), null);

                styleClass.AddDeclarationList(ft.GenerateDeclarationList(_context));

                RuleCollection.AddRule(styleClass);
            }
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
                    RuleCollection.AddRule(styleClass);
                }
            }
        }

        internal void AddToCollection(string name, ExcelTableStyleElement element, string htmlElement)
        {
            if (element.Style.HasValue == false) return; //Dont add empty elements

            var s = element.Style;

            var styleClass = new CssRule($"table.{name}{htmlElement}");

            var translators = new List<TranslatorBase>();

            if (s.Fill != null && _context.Exclude.Fill == false)
            {
                //TODO: Ensure if gradients with more than 2 colors it is handled correctly.
                translators.Add(new CssFillTranslator(new FillDxf(s.Fill)));
            }
            if (s.Font != null && _context.Exclude.Font != eFontExclude.All)
            {
                translators.Add(new CssFontTranslator(new FontDxf(s.Font), null));
            }
            if(s.Border != null && _context.Exclude.Border != eBorderExclude.All)
            {
                translators.Add(new CssBorderTranslator(new BorderDxf(s.Border)));
            }

            foreach (var translator in translators)
            {
                _context.SetTranslator(translator);
                _context.AddDeclarations(styleClass);
            }

            RuleCollection.AddRule(styleClass);
        }

        internal void AddToCollectionVH(string name, ExcelTableStyleElement element, string htmlElement)
        {
            if (element.Style.Border.Vertical.HasValue == false && element.Style.Border.Horizontal.HasValue == false) return; //Dont add empty elements

            var s = (IStyleExport)element.Style;

            var styleClass = new CssRule($"table.{name}{htmlElement}td,tr ");
            if (s.Border != null)
            {
                var translator = new CssBorderTranslator(s.Border);
                styleClass.AddDeclarationList(translator.GenerateDeclarationList(_context));
                RuleCollection.AddRule(styleClass);
            }
        }

        internal void AddOtherCollectionToThisCollection(CssRuleCollection otherCollection)
        {
            foreach (var otherRule in otherCollection)
            {
                RuleCollection.AddRule(otherRule);
            }
        }
    }
}
