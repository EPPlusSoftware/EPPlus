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
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Exporters
{
    internal abstract class AbstractHtmlExporter
    {
        public AbstractHtmlExporter()
        {
        }

        internal const string TableClass = "epplus-table";
        internal List<HtmlImage> _rangePictures = null;
        protected List<string> _dataTypes = new List<string>();
        protected readonly CellDataWriter _cellDataWriter = new CellDataWriter();
        protected Dictionary<string, int> _styleCache;

        internal void SetStyleCache(Dictionary<string, int> styleCache)
        {
            _styleCache = styleCache;
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

        internal void LoadRangeImages(List<ExcelRangeBase> ranges)
        {
            if (_rangePictures != null)
            {
                return;
            }
            _rangePictures = new List<HtmlImage>();
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

        protected string GetImageCellClassName(HtmlImage image, HtmlExportSettings settings)
        {
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
    }
}
