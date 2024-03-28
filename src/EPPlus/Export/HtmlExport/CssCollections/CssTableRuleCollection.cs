/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.Exporters.Internal;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using static OfficeOpenXml.Export.HtmlExport.ColumnDataTypeManager;

namespace OfficeOpenXml.Export.HtmlExport.CssCollections
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

            _context = new TranslatorContext(settings, _settings.Css.Exclude.TableStyle);
            _context.Theme = _theme;

            RuleCollection = new CssRuleCollection();
        }

        internal void AddHyperlink(string name, ExcelTableStyleElement element)
        {
            if (_context.Exclude.Font != eFontExclude.All &&
                element.Style.HasValue &&
                element.Style.Font.HasValue)
            {
                var styleClass = new CssRule($"table.{name} a", int.MaxValue);

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
                //if (_context.Exclude.HorizontalAlignment == false)
                //{
                //    rightDefault = false;
                //}
                //else
                //{
                    rightDefault = c < dataTypes.Count &&
                        (dataTypes[c] == HtmlDataTypes.Number || dataTypes[c] == HtmlDataTypes.DateTime);
                //}

                if (styleId > 0)
                {
                    var styleClass = new CssRule($"table.{name} td:nth-child({col})", int.MaxValue);

                    var xfs = new StyleXml(_table.WorkSheet.Workbook.Styles.CellXfs[styleId]);
                    var translator = new CssTableTextFormatTranslator(xfs, rightDefault);

                    styleClass.AddDeclarationList(translator.GenerateDeclarationList(_context));

                    //TODO: If we exclude both horizontal and vertical we can get a class with empty declaration list here...
                    RuleCollection.AddRule(styleClass);
                }
            }
        }

        internal void AddToCollection(string name, ExcelTableStyleElement element, string htmlElement)
        {
            if (element.Style.HasValue == false) return; //Dont add empty elements

            var s = element.Style;

            var styleClass = new CssRule($"table.{name}{htmlElement}",int.MaxValue);

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
            if (s.Border != null && _context.Exclude.Border != eBorderExclude.All)
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

            var s = new StyleDxf(element.Style);

            var styleClass = new CssRule($"table.{name}{htmlElement} td,tr ", int.MaxValue);
            if (s.Border != null)
            {
                var translator = new CssBorderTranslator(s.Border);
                styleClass.AddDeclarationList(translator.GenerateDeclarationList(_context));
                RuleCollection.AddRule(styleClass);
            }
        }

        internal void AddTableToCollection(ExcelTable table, List<string> datatypes, string tableClassPreset)
        {
            //if (settings.Minify == false) styleWriter.WriteLine();
            ExcelTableNamedStyle tblStyle;
            if (table.TableStyle == TableStyles.Custom)
            {
                tblStyle = table.WorkSheet.Workbook.Styles.TableStyles[table.StyleName].As.TableStyle;
            }
            else
            {
                var tmpNode = table.WorkSheet.Workbook.StylesXml.CreateElement("c:tableStyle");
                tblStyle = new ExcelTableNamedStyle(table.WorkSheet.Workbook.Styles.NameSpaceManager, tmpNode, table.WorkSheet.Workbook.Styles);
                tblStyle.SetFromTemplate(table.TableStyle);
            }

            var tableClass = $"{tableClassPreset}{HtmlExportTableUtil.GetClassName(tblStyle.Name, $"tablestyle{table.Id}")}";

            AddHyperlink($"{tableClass}", tblStyle.WholeTable);
            AddAlignment($"{tableClass}", datatypes);

            AddToCollection($"{tableClass}", tblStyle.WholeTable, "");
            AddToCollectionVH($"{tableClass}", tblStyle.WholeTable, "");

            //Header
            AddToCollection($"{tableClass}", tblStyle.HeaderRow, " thead");
            AddToCollectionVH($"{tableClass}", tblStyle.HeaderRow, "");

            AddToCollection($"{tableClass}", tblStyle.LastTotalCell, $" thead tr th:last-child)");
            AddToCollection($"{tableClass}", tblStyle.FirstHeaderCell, " thead tr th:first-child");

            //Total
            AddToCollection($"{tableClass}", tblStyle.TotalRow, " tfoot");
            AddToCollectionVH($"{tableClass}", tblStyle.TotalRow, "");
            AddToCollection($"{tableClass}", tblStyle.LastTotalCell, $" tfoot tr td:last-child)");
            AddToCollection($"{tableClass}", tblStyle.FirstTotalCell, " tfoot tr td:first-child");

            //Columns stripes
            var tableClassCS = $"{tableClass}-column-stripes";
            AddToCollection($"{tableClassCS}", tblStyle.FirstColumnStripe, $" tbody tr td:nth-child(odd)");
            AddToCollection($"{tableClassCS}", tblStyle.SecondColumnStripe, $" tbody tr td:nth-child(even)");

            //Row stripes
            var tableClassRS = $"{tableClass}-row-stripes";
            AddToCollection($"{tableClassRS}", tblStyle.FirstRowStripe, " tbody tr:nth-child(odd)");
            AddToCollection($"{tableClassRS}", tblStyle.SecondRowStripe, " tbody tr:nth-child(even)");

            //Last column
            var tableClassLC = $"{tableClass}-last-column";
            AddToCollection($"{tableClassLC}", tblStyle.LastColumn, $" tbody tr td:last-child");

            //First column
            var tableClassFC = $"{tableClass}-first-column";
            AddToCollection($"{tableClassFC}", tblStyle.FirstColumn, " tbody tr td:first-child");
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
