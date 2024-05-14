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
using OfficeOpenXml.ConditionalFormatting.Rules;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Export.HtmlExport.Determinator;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Data;
using System.Runtime.CompilerServices;

namespace OfficeOpenXml.Export.HtmlExport.Exporters.Internal
{
    internal abstract class CssExporterBase : AbstractHtmlExporter
    {
        internal HashSet<int> _addedToCssXsf = new HashSet<int>();
		internal HashSet<int> _addedToCssDxf = new HashSet<int>();

        internal static int OrderDefaultXsf = int.MaxValue -1;
        internal static int OrderDefaultDxf = 0;

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
                if (tableSettings == null || tableSettings.Css.IncludeCellStyles)
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
                        if (tableSettings == null)
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
                            if (sc.ShouldAdd || _addedToCssXsf.Contains(sc.Id) == false)
                            {
                                _addedToCssXsf.Add(sc.Id);
                                collection.AddToCollection(sc.GetStyleList(), ns, sc.Id, OrderDefaultXsf);
                            }
                        }
                    }
                }
                if(Settings.RenderConditionalFormattings)
                {
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

            if (sc.ShouldAddWithBorders(bottomStyleId, rightStyleId) || _addedToCssXsf.Contains(sc.Id) == false)
            {
                _addedToCssXsf.Add(sc.Id);
                collection.AddToCollection(sc.GetStyleList(), range.Worksheet.Workbook.Styles.GetNormalStyle(), sc.Id, OrderDefaultXsf);
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
                    switch(cf.Value.Type)
                    {
                        case eExcelConditionalFormattingRuleType.TwoColorScale:
                        case eExcelConditionalFormattingRuleType.ThreeColorScale:
                            break;
                        case eExcelConditionalFormattingRuleType.DataBar:
                            var hasBeenAddedToCache = _exporterContext._dxfStyleCache.IsAdded($"{cf.Value.Uid}", out int cfId);
                            var hasBeenAddedToCss = _addedToCssDxf.Contains(cfId);

                            if (hasBeenAddedToCache == false || hasBeenAddedToCss == false)
                            {
                                cssTranslator.AddDatabar((ExcelConditionalFormattingDataBar)cf.Value, OrderDefaultDxf + cf.Value.Priority, cfId);
                                _addedToCssDxf.Add(cfId);
                            }
                            break;
                        case eExcelConditionalFormattingRuleType.ThreeIconSet:
                            AddIconSetToCollection((ExcelConditionalFormattingThreeIconSet)cf.Value.As.ThreeIconSet, cf.Value, cssTranslator);
                            break;
                        case eExcelConditionalFormattingRuleType.FourIconSet:
                            AddIconSetToCollection((ExcelConditionalFormattingFourIconSet)cf.Value.As.FourIconSet, cf.Value, cssTranslator);
                            break;
                        case eExcelConditionalFormattingRuleType.FiveIconSet:
                            AddIconSetToCollection((ExcelConditionalFormattingFiveIconSet)cf.Value.As.FiveIconSet, cf.Value, cssTranslator);
                            break;
                        default:
                            if (cf.Value.Style.HasValue)
                            {
                                var style = new StyleDxf(cf.Value.Style);
                                if (!_exporterContext._dxfStyleCache.IsAdded(style.StyleKey, out int id) || _addedToCssDxf.Contains(id) == false)
                                {
                                    _addedToCssDxf.Add(id);
                                    var name = $".{Settings.StyleClassPrefix}{Settings.DxfStyleClassName}{id}";
                                    cssTranslator.AddToCollection(new List<IStyleExport>() { style }, normalStyle, id, OrderDefaultDxf + cf.Value.Priority, name);
                                }
                            }
                            break;
                    }
                }
            }
        }

        internal void AddIconSetToCollection<T>(ExcelConditionalFormattingIconSetBase<T> iconSet, ExcelConditionalFormattingRule rule, CssRangeRuleCollection cssTranslator) 
            where T : struct, Enum
        {
            var hasBeenAddedToCache = _exporterContext._dxfStyleCache.IsAdded($"{rule.Uid}", out int cfId);
            var hasBeenAddedToCss = _addedToCssDxf.Contains(cfId);

            if (hasBeenAddedToCache == false || hasBeenAddedToCss == false)
            {
                cssTranslator.AddIconSetCF(iconSet, OrderDefaultDxf + rule.Priority, cfId);
                _addedToCssDxf.Add(cfId);
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

        internal CssWriter GetTableCssWriter(Stream stream, ExcelTable table, HtmlTableExportSettings tableSettings)
        {
            if ((table.TableStyle == TableStyles.None || tableSettings.Css.IncludeTableStyles == false) && tableSettings.Css.IncludeCellStyles == false)
            {
                return null;
            }
            var cssWriter = new CssWriter(stream);

            if (_dataTypes.Count == 0) GetDataTypes(table.Address, table);
            return cssWriter;
        }
    }
}
