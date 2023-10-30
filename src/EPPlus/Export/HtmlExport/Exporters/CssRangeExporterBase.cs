using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Export.HtmlExport.Collectors;
using OfficeOpenXml.Export.HtmlExport.Determinator;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Input;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class CssRangeExporterBase : CssExporterBase
    {
        internal CssRangeExporterBase(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges) : base(settings, ranges) 
        {
            _settings = settings;
        }

        public CssRangeExporterBase(HtmlRangeExportSettings settings, ExcelRangeBase range)
            : base(settings, range)
        {
            _settings = settings;
        }

        protected readonly HtmlRangeExportSettings _settings;

        protected void RenderCellCss(StreamWriter sw)
        {
            var cssTranslator = new CssRangeRuleCollection(_ranges._list, _settings);
            var trueWriter = new CssTrueWriter(sw);

            cssTranslator.AddSharedClasses(TableClass);


            AddRangesToCollection(cssTranslator);

            if (Settings.Pictures.Include == ePictureInclude.Include)
            {
                LoadRangeImages(_ranges._list);
                foreach (var p in _rangePictures)
                {
                    cssTranslator.AddPictureToCss(p);
                }
            }

            WriteAndClearCollection(cssTranslator.RuleCollection, trueWriter);
            sw.Flush();
        }

        void AddRangesToCollection(CssRangeRuleCollection cssTranslator)
        {
            var addedTableStyles = new HashSet<TableStyles>();

            foreach (var range in _ranges._list)
            {

                var ws = range.Worksheet;
                var styles = ws.Workbook.Styles;
                var ce = new CellStoreEnumerator<ExcelValue>(range.Worksheet._values, range._fromRow, range._fromCol, range._toRow, range._toCol);
                ExcelAddressBase address = null;

                while (ce.Next())
                {
                    if (ce.Value._styleId > 0 && ce.Value._styleId < styles.CellXfs.Count)
                    {
                        var ma = ws.MergedCells[ce.Row, ce.Column];
                        var xfs = styles.CellXfs[ce.Value._styleId];

                        var sc = new StyleChecker(styles);
                        sc.Style = new StyleXml(xfs);
                        sc.Cache = _exporterContext._styleCache;

                        if (ma != null)
                        {
                            if (address == null || address.Address != ma)
                            {
                                address = new ExcelAddressBase(ma);
                            }
                            var fromRow = address._fromRow < range._fromRow ? range._fromRow : address._fromRow;
                            var fromCol = address._fromCol < range._fromCol ? range._fromCol : address._fromCol;

                            if (fromRow != ce.Row || fromCol != ce.Column) //Only add the style for the top-left cell in the merged range.
                                continue;

                            var mAdr = new ExcelAddressBase(ma);
                            var bottomStyleId = range.Worksheet._values.GetValue(mAdr._toRow, mAdr._fromCol)._styleId;
                            var rightStyleId = range.Worksheet._values.GetValue(mAdr._fromRow, mAdr._toCol)._styleId;

                            if (sc.ShouldAddWithBorders(bottomStyleId, rightStyleId))
                            {
                                cssTranslator.AddToCollection(sc.GetStyleList(), styles.GetNormalStyle(), sc.Id);
                            }

                            if (sc.ShouldAdd)
                            {
                                cssTranslator.AddToCollection(sc.GetStyleList(), styles.GetNormalStyle(), sc.Id);
                            }

                            //AddConditionalFormattingsToCollection(ce.CellAddress, styles.GetNormalStyle(), cssTranslator);

                            //if (ce.CellAddress != null)
                            //{
                            //    var items = GetCFItemsAtAddress(ce.CellAddress);

                            //    foreach (var cf in items)
                            //    {
                            //        var style = new StyleDxf(cf.Value.Style);
                            //        //TODO: Figure out how to determine bottom and right when merged cell cfs
                            //        if (!_exporterContext._dxfStyleCache.IsAdded(style.StyleKey, out int id))
                            //        {
                            //            AddStyleToCollection(id, new List<IStyleExport>() { style });
                            //        }
                            //    }
                            //}
                        }
                        else
                        {
                            if (sc.ShouldAdd)
                            {
                                cssTranslator.AddToCollection(sc.GetStyleList(), styles.GetNormalStyle(), sc.Id);
                            }
                        }

                        AddConditionalFormattingsToCollection(ce.CellAddress, styles.GetNormalStyle(), cssTranslator);
                    }
                }

                //if (Settings.TableStyle == eHtmlRangeTableInclude.Include)
                //{
                //    var table = range.GetTable();
                //    if (table != null &&
                //       table.TableStyle != TableStyles.None &&
                //       addedTableStyles.Contains(table.TableStyle) == false)
                //    {
                //        var settings = new HtmlTableExportSettings() { Minify = Settings.Minify };
                //        HtmlExportTableUtil.RenderTableCss(sw, table, settings, _exporterContext._styleCache, _dataTypes);
                //        addedTableStyles.Add(table.TableStyle);
                //    }
                //}
            }
        }

        internal void AddConditionalFormattingsToCollection(string cellAddress, ExcelNamedStyleXml normalStyle,  CssRangeRuleCollection cssTranslator)
        {
            if (cellAddress != null)
            {
                var items = GetCFItemsAtAddress(cellAddress);

                foreach (var cf in items)
                {
                    var style = new StyleDxf(cf.Value.Style);
                    if (!_exporterContext._dxfStyleCache.IsAdded(style.StyleKey, out int id))
                    {
                        cssTranslator.AddToCollection(new List<IStyleExport>() { style }, normalStyle, id);
                    }
                }
            }
        }

        internal List<QuadRangeItem<ExcelConditionalFormattingRule>> GetCFItemsAtAddress(string cellAddress)
        {
            return _cfQuadTree.GetIntersectingRangeItems
                (new QuadRange(new ExcelAddress(cellAddress)));
        }

        internal void WriteAndClearCollection(CssRuleCollection collection, CssTrueWriter writer)
        {
            for (int i = 0; i < collection.CssRules.Count(); i++)
            {
                writer.WriteRule(collection[i], _settings.Minify);
            }

            collection.CssRules.Clear();
        }


    }
}
