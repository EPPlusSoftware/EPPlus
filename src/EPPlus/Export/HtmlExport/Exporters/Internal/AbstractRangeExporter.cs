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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils.String;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.Export.HtmlExport.Exporters.Internal
{
    internal abstract class AbstractHtmlExporter
    {
        public AbstractHtmlExporter()
        {
        }

        internal const string TableClass = "epplus-table";
        internal List<HtmlImage> _rangePictures = null;
        internal List<HtmlSvgDrawing> _rangeDrawings = null;
        protected List<string> _dataTypes = new List<string>();
        protected ExporterContext _exporterContext;

        internal void SetExporterContext(ExporterContext context)
        {
            _exporterContext = context;
        }

        protected void GetDataTypes(ExcelAddressBase adr, ExcelTable table)
        {
            _dataTypes = new List<string>();
            for (int col = adr._fromCol; col <= adr._toCol; col++)
            {
                _dataTypes.Add(
                    ColumnDataTypeManager.GetColumnDataType(table.WorkSheet, table.Range, 2, col));
            }
        }

        internal void LoadRangeDrawings(List<ExcelRangeBase> ranges)
        {
            if (_rangePictures != null)
            {
                return;
            }
            _rangePictures = new List<HtmlImage>();
            _rangeDrawings = new List<HtmlSvgDrawing>();
            //Render in-cell images.
            foreach (var worksheet in ranges.Select(x => x.Worksheet).Distinct())
            {
                foreach (var d in worksheet.Drawings)
                {
                    if (d is ExcelPicture p)
                    {
                        p.GetFromBounds(out int fromRow, out int fromRowOff, out int fromCol, out int fromColOff);
                        p.GetToBounds(out int toRow, out int toRowOff, out int toCol, out int toColOff);

                        _rangePictures.Add(new HtmlImage()
                        {
                            WorksheetId = worksheet.PositionId,
                            Picture = p,
                            FromRow = fromRow,
                            FromRowOff = fromRowOff,
                            FromColumn = fromCol,
                            FromColumnOff = fromColOff,
                            ToRow = toRow,
                            ToRowOff = toRowOff,
                            ToColumn = toCol,
                            ToColumnOff = toColOff
                        });
                    }
                    else if(d.SupportsSvgExport && (d is ExcelShape || d is ExcelChart))
                    {
                        d.GetFromBounds(out int fromRow, out int fromRowOff, out int fromCol, out int fromColOff);
                        d.GetToBounds(out int toRow, out int toRowOff, out int toCol, out int toColOff);

                        _rangeDrawings.Add(new HtmlSvgDrawing()
                        {
                            WorksheetId = worksheet.PositionId,
                            Drawing = d,
                            FromRow = fromRow,
                            FromRowOff = fromRowOff,
                            FromColumn = fromCol,
                            FromColumnOff = fromColOff,
                            ToRow = toRow,
                            ToRowOff = toRowOff,
                            ToColumn = toCol,
                            ToColumnOff = toColOff
                        });
                    }
                }
            }
        }

        protected string GetCellText(ExcelRangeBase cell, HtmlExportSettings settings)
        {
            if (cell.IsRichText)
            {
                return cell.RichText.HtmlText;
            }
            else
            {
                return ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, settings.Culture);
            }
        }

        protected string GetImageCellClassName(HtmlImage image, HtmlExportSettings settings, bool isTable = false)
        {
            if (isTable)
            {
                return image == null ? "" : settings.StyleClassPrefix + "image-cell";
            }

            return image == null && settings.Pictures.Position != ePicturePosition.Absolute ? "" : settings.StyleClassPrefix + "image-cell";
        }

        protected HtmlImage GetImage(int worksheetId, int row, int col)
        {
            if (_rangePictures == null) return null;
            foreach (var p in _rangePictures)
            {
                if (p.FromRow == row - 1 && p.FromColumn == col - 1 && p.WorksheetId == worksheetId)
                {
                    return p;
                }
            }
            return null;
        }

        protected HtmlSvgDrawing GetDrawing(int worksheetId, int row, int col)
        {
            if (_rangeDrawings == null) return null;
            foreach (var d in _rangeDrawings)
            {
                if (d.FromRow == row - 1 && d.FromColumn == col - 1 && d.WorksheetId == worksheetId)
                {
                    return d;
                }
            }
            return null;
        }
        /// <summary>
        /// Adjust all drawings for the worksheets dimension and include any draings that are outside the dimension.
        /// </summary>
        /// <param name="ranges"></param>
        /// <param name="includeDrawings"></param>
        protected void AdjustRangeForDimensionAndDrawings(List<ExcelRangeBase> ranges, bool includeDrawings)
        {
            for(int i=0;i<ranges.Count;i++)
            {
                ranges[i] = AdjustRangeForDimensionAndDrawings(ranges[i], includeDrawings);
            }
        }
        protected ExcelRangeBase AdjustRangeForDimensionAndDrawings(ExcelRangeBase range, bool includeDrawings)
        {
            var newRange = range.DimensionAdjustedAddress;
            if (includeDrawings)
            {
                var drawMinRow = int.MaxValue;
                var drawMinCol = int.MaxValue;
                var drawMaxRow = int.MinValue;
                var drawMaxCol = int.MinValue;

                foreach (var d in range.Worksheet.Drawings)
                {
                    if (d.SupportsSvgExport && (d.IncludeInHtmlExport == eDrawingInclude.Include || d.IncludeInHtmlExport == eDrawingInclude.IncludeInHtmlOnly))
                    {
                        d.GetFromBounds(out int fromRow, out int fromRowOff, out int fromCol, out int fromColOff);
                        d.GetToBounds(out int toRow, out int toRowOff, out int toCol, out int toColOff);

                        //Bounds are 0 based;                        
                        fromRow++;
                        fromCol++;

                        if (fromRow > drawMinRow)
                        {
                            if (toRowOff > 0) toRow++;
                            if (toColOff > 0) toCol++;
                        }

                        if (range.Collide(fromRow, fromCol, toRow, toCol) != ExcelAddressBase.eAddressCollition.Inside)
                        {
                            if (fromRow < drawMinRow) drawMinRow = fromRow;
                            if (fromCol < drawMinCol) drawMinCol = fromCol;
                            if (toRow > drawMaxRow) drawMaxRow = toRow;
                            if (toCol > drawMaxCol) drawMaxCol = toCol;
                        }
                    }
                }

                if (newRange != null &&
                    newRange._fromRow > drawMinRow ||
                    newRange._fromCol > drawMinCol ||
                    newRange._toRow > drawMinRow ||
                    newRange._toCol > drawMinCol)
                {
                    return range.Worksheet.Cells[drawMinRow < newRange._fromRow ? Math.Max(drawMinRow, range._fromRow) : newRange._fromRow,
                                 drawMinCol < newRange._fromCol ? Math.Max(drawMinCol, range._fromCol) : newRange._fromCol,
                                 drawMaxRow > newRange._toRow ? Math.Min(drawMaxRow, range._toRow) : newRange._toRow,
                                 drawMaxCol > newRange._toCol ? Math.Min(drawMaxCol, range._toCol) : newRange._toCol];
                }
            }
            return range.Worksheet.Cells[newRange.Address];
        }
    }
}
