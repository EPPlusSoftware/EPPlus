/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  6/4/2022         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Export.HtmlExport.Collectors;
using OfficeOpenXml.Export.HtmlExport.Determinator;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class CssExporterBase : AbstractHtmlExporter
    {
        internal HashSet<int> _addedToCss = new HashSet<int>();

        public CssExporterBase(HtmlExportSettings settings, ExcelRangeBase range)
        {
            Settings = settings;
            Require.Argument(range).IsNotNull("range");
            _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            if (range.Addresses == null)
            {
                AddRange(range);
            }
            else
            {
                foreach (var address in range.Addresses)
                {
                    AddRange(range.Worksheet.Cells[address.Address]);
                }
            }
        }

        public CssExporterBase(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
        {
            Settings = settings;
            Require.Argument(ranges).IsNotNull("ranges");
            _ranges = ranges;
        }

        protected HtmlExportSettings Settings;
        protected EPPlusReadOnlyList<ExcelRangeBase> _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();
        internal const string TableStyleClassPrefix = "ts-";

        private void AddRange(ExcelRangeBase range)
        {
            if (range.IsFullColumn && range.IsFullRow)
            {
                _ranges.Add(new ExcelRangeBase(range.Worksheet, range.Worksheet.Dimension.Address));
            }
            else
            {
                _ranges.Add(range);
            }
        }

        protected CssRuleCollection CreateRuleCollection(HtmlRangeExportSettings settings)
        {
            var cssTranslator = new CssRangeRuleCollection(_ranges._list, settings);

            AddCssRulesToCollection(cssTranslator);

            return cssTranslator.RuleCollection;
        }

        protected CssRuleCollection CreateRuleCollection(HtmlTableExportSettings settings)
        {
            var cssTranslator = new CssRangeRuleCollection(_ranges._list, settings);

            AddCssRulesToCollection(cssTranslator, settings);

            return cssTranslator.RuleCollection;
        }

        protected void AddCssRulesToCollection(CssRangeRuleCollection cssTranslator, HtmlTableExportSettings tableSettings = null)
        {
            cssTranslator.AddSharedClasses(TableClass);

            var addedTableStyles = new HashSet<TableStyles>();

            foreach (var range in _ranges._list)
            {
                if(tableSettings == null || tableSettings.Css.IncludeCellStyles)
                {
                    AddCellCss(cssTranslator, range, tableSettings != null);
                }

                if (Settings.TableStyle == eHtmlRangeTableInclude.Include || tableSettings != null && tableSettings.Css.IncludeTableStyles)
                {
                    var table = range.GetTable();
                    if (table != null &&
                       table.TableStyle != TableStyles.None &&
                       addedTableStyles.Contains(table.TableStyle) == false)
                    {
                        if(tableSettings == null)
                        {
                            tableSettings = new HtmlTableExportSettings() { Minify = Settings.Minify };
                        }

                        cssTranslator.AddOtherCollectionToThisCollection
                            (
                                CreateTableCssRules(table, tableSettings, _dataTypes).RuleCollection
                            );
                        addedTableStyles.Add(table.TableStyle);
                    }
                }
            }

            if (Settings.Pictures.Include == ePictureInclude.Include)
            {
                LoadRangeImages(_ranges._list);
                foreach (var p in _rangePictures)
                {
                    cssTranslator.AddPictureToCss(p);
                }
            }
        }

        protected void AddCellCss(CssRangeRuleCollection collection, ExcelRangeBase range, bool isTableExporter = false)
        {
            var styles = range.Worksheet.Workbook.Styles;
            var ns = styles.GetNormalStyle();
            var ce = new CellStoreEnumerator<ExcelValue>(range.Worksheet._values, range._fromRow, range._fromCol, range._toRow, range._toCol);

            while (ce.Next())
            {
                if (ce.Value._styleId > 0 && ce.Value._styleId < styles.CellXfs.Count)
                {
                    var style = new StyleXml(styles.CellXfs[ce.Value._styleId]);
                    if (style.HasStyle)
                    {
                        var sc = new StyleChecker(styles, style, _exporterContext._styleCache);
                        var ma = range.Worksheet.MergedCells[ce.Row, ce.Column];

                        if (!isTableExporter && ma != null)
                        {
                            if (!AddMergedCellsToCollection(range, ma, ce, sc, collection))
                            {
                                continue;
                            }
                        }
                        else
                        {
                            if (sc.ShouldAdd || _addedToCss.Contains(sc.Id) == false)
                            {
                                _addedToCss.Add(sc.Id);
                                collection.AddToCollection(sc.GetStyleList(), ns, sc.Id);
                            }
                        }
                    }

                    AddConditionalFormattingsToCollection(ce.CellAddress, ns, collection);
                }
            }
        }

        private bool AddMergedCellsToCollection(ExcelRangeBase range, string ma, CellStoreEnumerator<ExcelValue> ce, StyleChecker sc, CssRangeRuleCollection collection)
        {
            var address = new ExcelAddressBase(ma);

            var fromRow = address._fromRow < range._fromRow ? range._fromRow : address._fromRow;
            var fromCol = address._fromCol < range._fromCol ? range._fromCol : address._fromCol;

            if (fromRow != ce.Row || fromCol != ce.Column) //Only add the style for the top-left cell in the merged range.
                return false;

            var mAdr = new ExcelAddressBase(ma);
            var bottomStyleId = range.Worksheet._values.GetValue(mAdr._toRow, mAdr._fromCol)._styleId;
            var rightStyleId = range.Worksheet._values.GetValue(mAdr._fromRow, mAdr._toCol)._styleId;

            if (sc.ShouldAddWithBorders(bottomStyleId, rightStyleId) || _addedToCss.Contains(sc.Id) == false)
            {
                _addedToCss.Add(sc.Id);
                collection.AddToCollection(sc.GetStyleList(), range.Worksheet.Workbook.Styles.GetNormalStyle(), sc.Id);
            }

            return true;
        }


        internal void AddConditionalFormattingsToCollection(string cellAddress, ExcelNamedStyleXml normalStyle, CssRangeRuleCollection cssTranslator)
        {
            if (cellAddress != null)
            {
                var items = GetCFItemsAtAddress(cellAddress);

                foreach (var cf in items)
                {
                    if(cf.Value.Style.HasValue)
                    {
                        var style = new StyleDxf(cf.Value.Style);
                        if (!_exporterContext._dxfStyleCache.IsAdded(style.StyleKey, out int id) || _addedToCss.Contains(id) == false)
                        {
                            _addedToCss.Add(id);
                            var name = $".{Settings.StyleClassPrefix}{Settings.CellStyleClassName}-dxf.id{id}";
                            cssTranslator.AddToCollection(new List<IStyleExport>() { style }, normalStyle, id, name);
                        }
                    }
                }
            }
        }

        internal List<QuadRangeItem<ExcelConditionalFormattingRule>> GetCFItemsAtAddress(string cellAddress)
        {
            return _exporterContext._cfQuadTree.GetIntersectingRangeItems
                (new QuadRange(new ExcelAddress(cellAddress)));
        }


        internal static CssTableRuleCollection CreateTableCssRules(ExcelTable table, HtmlTableExportSettings settings, List<string> datatypes)
        {
            var tableRules = new CssTableRuleCollection(table, settings);
            var tableClass = $"{TableClass}.{TableStyleClassPrefix}";
            tableRules.AddTableToCollection(table, datatypes, tableClass);

            return tableRules;
        }
    }
}
