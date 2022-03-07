/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Table;
using System.IO;
using OfficeOpenXml.Utils;
#if !NET35 && !NET40
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Exports a <see cref="ExcelTable"/> to Html
    /// </summary>
    public partial class ExcelHtmlRangeExporter
    {        
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

            if (_datatypes.Count == 0) GetDataTypes();
            var sw = new StreamWriter(stream);
            await RenderCellCssAsync(sw);
        }

        private async Task RenderCellCssAsync(StreamWriter sw)
        {            
            var styleWriter = new EpplusCssWriter(sw, _range, Settings, Settings.Css, Settings.Css.CssExclude);
            
            await styleWriter.RenderAdditionalAndFontCssAsync(TableClass);
            var ws = _range.Worksheet;
            var styles = ws.Workbook.Styles;
            var ce = new CellStoreEnumerator<ExcelValue>(_range.Worksheet._values, _range._fromRow, _range._fromCol, _range._toRow, _range._toCol);
            ExcelAddressBase address = null;
            while (ce.Next())
            {
                if (ce.Value._styleId > 0 && ce.Value._styleId < styles.CellXfs.Count)
                {
                    var ma = ws.MergedCells[ce.Row, ce.Column];
                    if(ma!=null)
                    {
                        if (address == null || address.Address != ma)
                        {
                            address = new ExcelAddressBase(ma);
                        }
                        var fromRow = address._fromRow < _range._fromRow ? _range._fromRow : address._fromRow;
                        var fromCol = address._fromCol < _range._fromCol ? _range._fromCol : address._fromCol;
                        if (fromRow != ce.Row || fromCol != ce.Column) //Only add the style for the top-left cell in the merged range.
                            continue;                        
                    }
                    await styleWriter.AddToCssAsync(styles, ce.Value._styleId, Settings.StyleClassPrefix);
                }
            }
            if (Settings.Pictures.Include == ePictureInclude.Include)
            {
                LoadRangeImages(_range);
                foreach (var p in _rangePictures)
                {
                    await styleWriter.AddPictureToCssAsync(p);
                }
            }
            await styleWriter.FlushStreamAsync();
        }
    }
}
#endif
