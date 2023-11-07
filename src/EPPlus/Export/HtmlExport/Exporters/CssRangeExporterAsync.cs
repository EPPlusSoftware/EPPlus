﻿/*************************************************************************************************
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
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class CssRangeExporterAsync : CssRangeExporterBase
    {
        public CssRangeExporterAsync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
         : base(settings, ranges)
        {
            _settings = settings;
        }

        public CssRangeExporterAsync(HtmlRangeExportSettings settings, ExcelRangeBase range)
            : base(settings, range)
        {
            _settings = settings;
        }

        private readonly HtmlRangeExportSettings _settings;

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public async Task<string> GetCssStringAsync()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                await RenderCssAsync(ms);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return await sr.ReadToEndAsync();
                }
            }
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to html and writes it to a stream
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns></returns>
        public async Task RenderCssAsync(Stream stream)
        {
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writable System.IO.Stream");
            }

            //if (_datatypes.Count == 0) GetDataTypes();
            var sw = new StreamWriter(stream);
            await WriteCellAsync(sw);
        }

        private async Task WriteCellAsync(StreamWriter sw)
        {
            var trueWriter = new CssTrueWriter(sw);
            var cssTranslator = RenderCellCss(sw);

            await WriteAndClearCollection(cssTranslator.RuleCollection, trueWriter);
            sw.Flush();
        }

        private async Task WriteAndClearCollection(CssRuleCollection collection, CssTrueWriter writer)
        {
            for (int i = 0; i < collection.CssRules.Count(); i++)
            {
                await writer.WriteRuleAsync(collection[i], _settings.Minify);
            }

            collection.CssRules.Clear();
        }

        //private async Task RenderCellCssAsync(StreamWriter sw)
        //{
        //    var styleWriter = new EpplusCssWriter(sw, _ranges._list, _settings, _settings.Css, _settings.Css.CssExclude);

        //    await styleWriter.RenderAdditionalAndFontCssAsync(TableClass);
        //    var addedTableStyles = new HashSet<TableStyles>();
        //    foreach (var range in _ranges._list)
        //    {
        //        var ws = range.Worksheet;
        //        var styles = ws.Workbook.Styles;
        //        var ce = new CellStoreEnumerator<ExcelValue>(range.Worksheet._values, range._fromRow, range._fromCol, range._toRow, range._toCol);
        //        ExcelAddressBase address = null;
        //        while (ce.Next())
        //        {
        //            if (ce.Value._styleId > 0 && ce.Value._styleId < styles.CellXfs.Count)
        //            {
        //                var ma = ws.MergedCells[ce.Row, ce.Column];
        //                if (ma != null)
        //                {
        //                    if (address == null || address.Address != ma)
        //                    {
        //                        address = new ExcelAddressBase(ma);
        //                    }
        //                    var fromRow = address._fromRow < range._fromRow ? range._fromRow : address._fromRow;
        //                    var fromCol = address._fromCol < range._fromCol ? range._fromCol : address._fromCol;

        //                    if (fromRow != ce.Row || fromCol != ce.Column) //Only add the style for the top-left cell in the merged range.
        //                        continue;

        //                    var mAdr = new ExcelAddressBase(ma);
        //                    var bottomStyleId = range.Worksheet._values.GetValue(mAdr._toRow, mAdr._fromCol)._styleId;
        //                    var rightStyleId = range.Worksheet._values.GetValue(mAdr._fromRow, mAdr._toCol)._styleId;


        //                    var stylesList = new List<ExcelXfs>
        //                    {
        //                        styles.CellXfs[ce.Value._styleId],
        //                        styles.CellXfs[bottomStyleId],
        //                        styles.CellXfs[rightStyleId]
        //                    };

        //                    if (!StyleToCss.IsAddedToCache(stylesList[0], _exporterContext._dxfStyleCache, out int id))
        //                    {
        //                        if (AttributeTranslator.HasStyle(stylesList[0]))
        //                            styleWriter.AddToCss(stylesList, styles.GetNormalStyle(), Settings.StyleClassPrefix, Settings.CellStyleClassName, id);
        //                    }
        //                }
        //                else
        //                {
        //                    var xfs = styles.CellXfs[ce.Value._styleId];
        //                    if (!StyleToCss.IsAddedToCache(xfs, _exporterContext._styleCache, out int id))
        //                    {
        //                        if (AttributeTranslator.HasStyle(xfs))
        //                            await styleWriter.AddToCssAsync(xfs, styles.GetNormalStyle(), Settings.StyleClassPrefix, Settings.CellStyleClassName, id);
        //                    }

        //                    if (ce.CellAddress != null)
        //                    {
        //                        var items = GetCFItemsAtAddress(ce.CellAddress);

        //                        foreach (var cf in items)
        //                        {
        //                            var idDxf = StyleToCss.GetIdFromCache(cf.Value._style, _exporterContext._dxfStyleCache);
        //                            if(idDxf != -1)
        //                            {
        //                                await styleWriter.AddToCssAsyncCF(cf.Value._style, Settings.StyleClassPrefix, Settings.CellStyleClassName, idDxf);
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        if (Settings.TableStyle == eHtmlRangeTableInclude.Include)
        //        {
        //            var table = range.GetTable();
        //            if (table != null &&
        //               table.TableStyle != TableStyles.None &&
        //               addedTableStyles.Contains(table.TableStyle) == false)
        //            {
        //                var settings = new HtmlTableExportSettings() { Minify = Settings.Minify };
        //                await HtmlExportTableUtil.RenderTableCssAsync(sw, table, settings, _exporterContext._styleCache, _dataTypes);
        //                addedTableStyles.Add(table.TableStyle);
        //            }
        //        }
        //    }
        //    if (Settings.Pictures.Include == ePictureInclude.Include)
        //    {
        //        LoadRangeImages(_ranges._list);
        //        foreach (var p in _rangePictures)
        //        {
        //            await styleWriter.AddPictureToCssAsync(p);
        //        }
        //    }
        //    await styleWriter.FlushStreamAsync();
        //}

        //internal List<QuadRangeItem<ExcelConditionalFormattingRule>> GetCFItemsAtAddress(string cellAddress)
        //{
        //    return _exporterContext._cfQuadTree.GetIntersectingRangeItems
        //        (new QuadRange(new ExcelAddress(cellAddress)));
        //}
    }
}
#endif
