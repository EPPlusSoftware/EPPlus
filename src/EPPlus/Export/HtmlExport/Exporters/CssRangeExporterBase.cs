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
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using System;
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

        public static object TableStyleClassPrefix { get; private set; }

        protected CssRangeRuleCollection RenderCellCss(StreamWriter sw)
        {
            var cssTranslator = new CssRangeRuleCollection(_ranges._list, _settings);

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

            return cssTranslator;
        }

        protected void AddRangesToCollection(CssRangeRuleCollection cssTranslator)
        {
            var addedTableStyles = new HashSet<TableStyles>();

            foreach (var range in _ranges._list)
            {
                var ws = range.Worksheet;
                var styles = ws.Workbook.Styles;
                var ns = styles.GetNormalStyle();
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
                                cssTranslator.AddToCollection(sc.GetStyleList(), ns, sc.Id);
                            }
                        }
                        else
                        {
                            if (sc.ShouldAdd)
                            {
                                cssTranslator.AddToCollection(sc.GetStyleList(), ns, sc.Id);
                            }
                        }

                        AddConditionalFormattingsToCollection(ce.CellAddress, ns, cssTranslator);
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
                //        RenderTableCss(sw, table, settings, _exporterContext._styleCache, _dataTypes);
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
            return _exporterContext._cfQuadTree.GetIntersectingRangeItems
                (new QuadRange(new ExcelAddress(cellAddress)));
        }

        internal static void RenderTableCss(StreamWriter sw, ExcelTable table, HtmlTableExportSettings settings, List<string> datatypes)
        {
            var styleWriter = new EpplusTableCssWriter(sw, table, settings);
            if (settings.Minify == false) styleWriter.WriteLine();
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

            var tableClass = $"{TableClass}.{TableStyleClassPrefix}{HtmlExportTableUtil.GetClassName(tblStyle.Name, "EmptyTableStyle").ToLower()}";
            styleWriter.AddHyperlinkCss($"{tableClass}", tblStyle.WholeTable);
            styleWriter.AddAlignmentToCss($"{tableClass}", datatypes);

            styleWriter.AddToCss($"{tableClass}", tblStyle.WholeTable, "");
            styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.WholeTable, "");

            //Header
            styleWriter.AddToCss($"{tableClass}", tblStyle.HeaderRow, " thead");
            styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.HeaderRow, "");

            styleWriter.AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" thead tr th:last-child)");
            styleWriter.AddToCss($"{tableClass}", tblStyle.FirstHeaderCell, " thead tr th:first-child");

            //Total
            styleWriter.AddToCss($"{tableClass}", tblStyle.TotalRow, " tfoot");
            styleWriter.AddToCssBorderVH($"{tableClass}", tblStyle.TotalRow, "");
            styleWriter.AddToCss($"{tableClass}", tblStyle.LastTotalCell, $" tfoot tr td:last-child)");
            styleWriter.AddToCss($"{tableClass}", tblStyle.FirstTotalCell, " tfoot tr td:first-child");

            //Columns stripes
            var tableClassCS = $"{tableClass}-column-stripes";
            styleWriter.AddToCss($"{tableClassCS}", tblStyle.FirstColumnStripe, $" tbody tr td:nth-child(odd)");
            styleWriter.AddToCss($"{tableClassCS}", tblStyle.SecondColumnStripe, $" tbody tr td:nth-child(even)");

            //Row stripes
            var tableClassRS = $"{tableClass}-row-stripes";
            styleWriter.AddToCss($"{tableClassRS}", tblStyle.FirstRowStripe, " tbody tr:nth-child(odd)");
            styleWriter.AddToCss($"{tableClassRS}", tblStyle.SecondRowStripe, " tbody tr:nth-child(even)");

            //Last column
            var tableClassLC = $"{tableClass}-last-column";
            styleWriter.AddToCss($"{tableClassLC}", tblStyle.LastColumn, $" tbody tr td:last-child");

            //First column
            var tableClassFC = $"{tableClass}-first-column";
            styleWriter.AddToCss($"{tableClassFC}", tblStyle.FirstColumn, " tbody tr td:first-child");

            styleWriter.FlushStream();
        }
    }
}
