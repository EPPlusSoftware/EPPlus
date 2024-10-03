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
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    public partial class ExcelRangeBase
    {
        /// <summary>
        /// Returns the in-cell image in the top-left cell of the range.
        /// </summary>
        /// <returns>An instance of <see cref="ExcelCellPicture"/> or null if no such image exists.</returns>
        internal ExcelCellPicture GetCellPicture()
        {
            return _worksheet._cellPicturesManager.GetCellPicture(_fromRow, _fromCol);
        }

        /// <summary>
        /// Adds an image to the top-left cell in the range.
        /// </summary>
        /// <param name="imageBytes">The image bytes</param>
        /// <param name="altText">Alt-text of the image</param>
        /// <param name="markAsDecorative">Set to true if the picture should be marked as decorative (for accessability).</param>
        public void SetCellPicture(byte[] imageBytes, string altText = null, bool markAsDecorative = false)
        {
            var calcOrigin = markAsDecorative ? CalcOrigins.StandaloneDecorative : CalcOrigins.StandAlone;
            _worksheet._cellPicturesManager.SetCellPicture(_fromRow, _fromCol, imageBytes, altText, calcOrigin);
        }
        /// <summary>
        /// Adds an image to the top-left cell in the range.
        /// </summary>
        /// <param name="path">File system path to the image file</param>
        /// <param name="altText">Alt-text of the image</param>
        ///  /// <param name="markAsDecorative">Set to true if the picture should be marked as decorative (for accessability).</param>
        internal void SetCellPicture(string path, string altText = null, bool markAsDecorative = false)
        {
            var imageBytes = File.ReadAllBytes(path);
            SetCellPicture(imageBytes, altText, markAsDecorative);
        }

        /// <summary>
        /// Adds an image to the top-left cell in the range.
        /// </summary>
        /// <param name="imageStream">A <see cref="System.IO.Stream" containing the image bytes/></param>
        /// <param name="altText">Alt-text of the image</param>
        /// <param name="markAsDecorative">Set to true if the picture should be marked as decorative (for accessability).</param>
        public void SetCellPicture(Stream imageStream, string altText = null, bool markAsDecorative = false)
        {
            var calcOrigin = markAsDecorative ? CalcOrigins.StandaloneDecorative : CalcOrigins.StandAlone;
            _worksheet._cellPicturesManager.SetCellPicture(_fromRow, _fromCol, imageStream, altText, calcOrigin);
        }

        /// <summary>
        /// Adds an image to the top-left cell in the range.
        /// </summary>
        /// <param name="image"></param>
        /// <param name="altText"></param>
        /// <param name="markAsDecorative">Set to true if the picture should be marked as decorative (for accessability).</param>
        public void SetCellPicture(ExcelImage image, string altText = null, bool markAsDecorative = false)
        {
            var calcOrigin = markAsDecorative ? CalcOrigins.StandaloneDecorative : CalcOrigins.StandAlone;
            _worksheet._cellPicturesManager.SetCellPicture(_fromRow, _fromCol, image, altText, calcOrigin);
        }
    }
}
