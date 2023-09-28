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
using OfficeOpenXml.Export.HtmlExport.Parsers;
using OfficeOpenXml.Export.HtmlExport.Settings;
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
            var styleWriter = new EpplusCssWriter(sw, _ranges._list, _settings, _settings.Css, _settings.Css.CssExclude);

            var cssTranslator = new CssRangeTranslator(_ranges._list, _settings);
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

                            var stylesList = new List<ExcelXfs>
                            {
                                styles.CellXfs[ce.Value._styleId],
                                styles.CellXfs[bottomStyleId],
                                styles.CellXfs[rightStyleId]
                            };

                            if (!StyleToCss.IsAddedToCache(stylesList[0], _exporterContext._dxfStyleCache, out int id))
                            {
                                if (AttributeTranslator.HasStyle(stylesList[0]) || stylesList[1].BorderId > 0 || stylesList[2].BorderId > 0)
                                {
                                    cssTranslator.AddToCollection(stylesList, styles.GetNormalStyle(), id);
                                }
                            }
                        }
                        else
                        {
                            var xfs = styles.CellXfs[ce.Value._styleId];

                            if (!StyleToCss.IsAddedToCache(xfs, _exporterContext._dxfStyleCache, out int id))
                            {
                                if (AttributeTranslator.HasStyle(xfs))
                                {
                                    cssTranslator.AddToCollection(xfs, styles.GetNormalStyle(), id);
                                }
                            }
                        }
                    }
                }

                WriteAndClearCollection(cssTranslator.RuleCollection, trueWriter);

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
                    styleWriter.AddPictureToCss(p);
                }
            }
            styleWriter.FlushStream();
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
