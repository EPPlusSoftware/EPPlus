/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/11/2024         EPPlus Software AB       Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml.Drawing.EMF;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.Structures.Constants;
using System.IO;

namespace OfficeOpenXml.CellPictures
{
    /// <summary>
    /// Pictures in the cell/range
    /// </summary>
    public class ExcelRangePicture
    {

        /// <summary>
        /// Constructur
        /// </summary>
        internal ExcelRangePicture(ExcelRangeBase range)
        {
            _range = range;
            _sheet = range.Worksheet;
            _cellPicturesManager = new CellPicturesManager(_sheet);
        }

        private readonly ExcelRangeBase _range;
        private readonly ExcelWorksheet _sheet;
        private readonly CellPicturesManager _cellPicturesManager;


        /// <summary>
        /// Returns true if the range has a picture, otherwise false
        /// </summary>
        public bool Exists
        {
            get
            {
                if(_cellPicturesManager.GetCellPicture(_range._fromRow, _range._fromCol) != null)
                {
                    return true;
                }
                return _cellPicturesManager.GetCellPicture(_range._fromRow, _range._fromCol, StructureTypes.WebImage) != null;
            }
        }

        /// <summary>
        /// Returns a picture of the top-left cell in the range.
        /// </summary>
        /// <returns>An <see cref="ExcelCellPicture"/> or null if it doesn't exist</returns>
        public ExcelCellPicture Get()
        {
            var v = _sheet.Cells[_range._fromRow, _range._fromCol].Value;
            var rdr = v != null ? v as RichDataReferenceValueError : default(RichDataReferenceValueError);
            if (rdr != null && (rdr.ReferenceType == RichDataReferenceTypes.LocalImage || rdr.ReferenceType == RichDataReferenceTypes.WebImage))
            {
                return rdr as ExcelCellPicture;
            }
            return null;
        }

        /// <summary>
        /// Adds (or replaces) a cell picture.
        /// EPPlus supports the following image types: Png, Jpg, Gif, Bmp, WebP, Tif, Ico
        /// </summary>
        /// <param name="imageBytes">byte array of the image file</param>
        /// 
        /// 
        /// <param name="altText">Alt text for the cell/range picture</param>
        /// <param name="isDecorative">Sets the decorative property (used by accessibility tools)</param>
        public void Set(byte[] imageBytes, string altText = null, bool isDecorative = false)
        {
            var calcOrigin = isDecorative ? CalcOrigins.StandaloneDecorative : CalcOrigins.StandAlone;
            var fromRow = _range._fromRow;
            var toRow = _range._toRow;
            var fromCol = _range._fromCol;
            var toCol = _range._toCol;
            for (var row = fromRow; row <= toRow; row++)
            {
                for (var col = fromCol; col <= toCol; col++)
                {
                    _cellPicturesManager.SetCellPicture(row, col, imageBytes, altText, calcOrigin);
                }
            }      
        }

        /// <summary>
        /// Adds (or replaces) a cell picture.
        /// EPPlus supports the following image types: Png, Jpg, Gif, Bmp, WebP, Tif, Ico
        /// </summary>
        /// <param name="imageStream"><see cref="System.IO.Stream"/> containing the image bytes</param>
        /// <param name="altText">Alt text for the cell/range picture</param>
        /// <param name="isDecorative">>Sets the decorative property (used by accessibility tools)</param>
        public void Set(Stream imageStream, string altText = null, bool isDecorative = false)
        {
            var calcOrigin = isDecorative ? CalcOrigins.StandaloneDecorative : CalcOrigins.StandAlone;
            var fromRow = _range._fromRow;
            var toRow = _range._toRow;
            var fromCol = _range._fromCol;
            var toCol = _range._toCol;
            for (var row = fromRow; row <= toRow; row++)
            {
                for (var col = fromCol; col <= toCol; col++)
                {
                    _cellPicturesManager.SetCellPicture(_range._fromRow, _range._fromCol, imageStream, altText, calcOrigin);
                }
            }   
        }

        /// <summary>
        /// Adds (or replaces) a cell picture.
        /// EPPlus supports the following image types: Png, Jpg, Gif, Bmp, WebP, Tif, Ico
        /// </summary>
        /// <param name="path">File path to the image file</param>
        /// <param name="altText">Alt text for the cell/range picture</param>
        /// <param name="isDecorative">>Sets the decorative property (used by accessibility tools)</param>
        public void Set(string path, string altText = null, bool isDecorative = false)
        {
            var calcOrigin = isDecorative ? CalcOrigins.StandaloneDecorative : CalcOrigins.StandAlone;
            var fromRow = _range._fromRow;
            var toRow = _range._toRow;
            var fromCol = _range._fromCol;
            var toCol = _range._toCol;
            for (var row = fromRow; row <= toRow; row++)
            {
                for (var col = fromCol; col <= toCol; col++)
                {
                    _cellPicturesManager.SetCellPicture(row, col, path, altText, calcOrigin);
                }
            }       
        }

        /// <summary>
        /// Adds (or replaces) a cell picture.
        /// EPPlus supports the following image types: Png, Jpg, Gif, Bmp, WebP, Tif, Ico
        /// </summary>
        /// <param name="fileInfo"><see cref="FileInfo" /> representing the path to the image file</param>
        /// <param name="altText">Alt text for the cell/range picture</param>
        /// <param name="isDecorative">>Sets the decorative property (used by accessibility tools)</param>
        public void Set(FileInfo fileInfo, string altText = null, bool isDecorative = false)
        {
            var calcOrigin = isDecorative ? CalcOrigins.StandaloneDecorative : CalcOrigins.StandAlone;
            var fromRow = _range._fromRow;
            var toRow = _range._toRow;
            var fromCol = _range._fromCol;
            var toCol = _range._toCol;
            for (var row = fromRow; row <= toRow; row++)
            {
                for(var col = fromCol; col <= toCol; col++)
                {
                    _cellPicturesManager.SetCellPicture(row, col, fileInfo, altText, calcOrigin);
                }
            }
        }

        /// <summary>
        /// Remove any cell picture in the entire range
        /// </summary>
        public void Remove()
        {
            var ws = _range.Worksheet;
            var maxRow = _range._toRow > ws.Dimension._toRow ? ws.Dimension._toRow : _range._toRow;
            var maxCol = _range._toCol > ws.Dimension._toCol ? ws.Dimension._toCol : _range._toCol;
            var range = ws.Cells[_range._fromRow, _range._fromCol, maxRow, maxCol];
            foreach(var cell in range)
            {
                _cellPicturesManager.RemoveCellPicture(cell._fromRow, cell._toCol);
            }
        }
    }
}
