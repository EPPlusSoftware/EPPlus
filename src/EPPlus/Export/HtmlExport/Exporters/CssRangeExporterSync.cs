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
using OfficeOpenXml.Core;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Export.HtmlExport.Collectors;
using OfficeOpenXml.Export.HtmlExport.Determinator;
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Export.HtmlExport.Settings;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal class CssRangeExporterSync : CssRangeExporterBase
    {
        public CssRangeExporterSync(HtmlRangeExportSettings settings, EPPlusReadOnlyList<ExcelRangeBase> ranges)
            : base(settings, ranges)
        {
            _settings = settings;
        }

        public CssRangeExporterSync(HtmlRangeExportSettings settings, ExcelRangeBase range)
            : base(settings, range)
        {
            _settings = settings;
        }

        private readonly HtmlRangeExportSettings _settings;

        //public HtmlRangeExportSettings Settings => _settings;

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetCssString()
        {
            using (var ms = RecyclableMemory.GetStream())
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
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writable System.IO.Stream");
            }

            //if (_datatypes.Count == 0) GetDataTypes();
            var sw = new StreamWriter(stream);
            RenderCellCss(sw);
        }

        private void RenderCellCss(StreamWriter sw)
        {
            var cssTranslator = new CssRangeRuleCollection(_ranges._list, _settings);
            var trueWriter = new CssTrueWriter(sw);

            cssTranslator.AddSharedClasses(TableClass);

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
                        var sc = new StyleChecker(xfs, _exporterContext._styleCache, styles);

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

                            if(sc.ShouldAddWithBorders(bottomStyleId, rightStyleId))
                            {
                              // cssTranslator.AddToCollection(sc.GetStyleList(), styles.GetNormalStyle(), sc.Id);
                            }
                        }
                        else
                        {
                            if (sc.ShouldAdd())
                            {
                                var style = new StyleXml(sc.GetStyleList()[0]);

                                cssTranslator.AddToCollection(new List<IStyle>() { style }, styles.GetNormalStyle(), sc.Id);
                            }

                            if (ce.CellAddress != null)
                            {
                                if (_cfAtAddresses.ContainsKey(ce.CellAddress))
                                {
                                    foreach (var cf in _cfAtAddresses[ce.CellAddress])
                                    {
                                        var idDxf = StyleToCss.GetIdFromCache(cf._style, _exporterContext._dxfStyleCache);
                                        if (idDxf != -1)
                                        {
                                            //var scDxf = new StyleChecker(cf.Style, _exporterContext._dxfStyleCache, styles);

                                            var dxfStyle = new StyleDxf(cf._style);

                                            //cf._style.Font.Color
                                            cssTranslator.AddToCollection(new List<IStyle>() { dxfStyle }, styles.GetNormalStyle(), idDxf);

                                            //await styleWriter.AddToCssAsyncCF(cf._style, Settings.StyleClassPrefix, Settings.CellStyleClassName, idDxf);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                if (Settings.TableStyle == eHtmlRangeTableInclude.Include)
                {
                    var table = range.GetTable();
                    if (table != null &&
                       table.TableStyle != TableStyles.None &&
                       addedTableStyles.Contains(table.TableStyle) == false)
                    {
                        var settings = new HtmlTableExportSettings() { Minify = Settings.Minify };
                        HtmlExportTableUtil.RenderTableCss(sw, table, settings, _exporterContext._styleCache, _dataTypes);
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

            WriteAndClearCollection(cssTranslator.RuleCollection, trueWriter);
            sw.Flush();
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
