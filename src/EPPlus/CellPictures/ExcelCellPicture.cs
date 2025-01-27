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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.RichValues.WebImages;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeOpenXml.CellPictures
{
    /// <summary>
    /// Represents an in-cell picture
    /// </summary>
    [DebuggerDisplay("Type: {PictureType}, Filename: {FileName}")]
    public class ExcelCellPicture : RichDataReferenceValueError
    {
        /// <summary>
        /// Constructor
        /// </summary>
        internal ExcelCellPicture(uint vmId, Uri imageUri, PictureStore pictureStore, ExcelCellPictureTypes pictureType) : base(vmId, PictureTypeToReferenceType(pictureType))
        {
            _pictureStore = pictureStore;
            ImageUri = imageUri;
            PictureType = pictureType;
        }

        private readonly PictureStore _pictureStore;

        private static RichDataReferenceTypes PictureTypeToReferenceType(ExcelCellPictureTypes pictureType)
        {
            if(pictureType == ExcelCellPictureTypes.WebImage)
            {
                return RichDataReferenceTypes.WebImage;
            }
            return RichDataReferenceTypes.LocalImage;
        }

        /// <summary>
        /// Internal uri in the workbook of the image.
        /// </summary>
        internal Uri ImageUri
        {
            get; set;
        }

        /// <summary>
        /// External Uri, only set for images retrieved via the IMAGE function
        /// </summary>
        public Uri ExternalAddress
        {
            get; internal set;
        }

        /// <summary>
        /// Type of cell picture
        /// </summary>
        public ExcelCellPictureTypes PictureType
        {
            get;
            private set;
        }

        /// <summary>
        /// The bytes of the image file
        /// </summary>
        /// <returns></returns>
        public byte[] GetImageBytes()
        {
            return _pictureStore.GetImageBytes(ImageUri);
        }

        /// <summary>
        /// Name of the image file including file extension
        /// </summary>
        public string FileName
        {
            get
            {
                if(ImageUri == null)
                {
                    return null;
                }
                return Path.GetFileName(ImageUri.OriginalString);
            }
        }

        /// <summary>
        /// Alt text of the image
        /// </summary>
        public string AltText
        {
            get; set;
        }

        /// <summary>
        /// Indicates the calculation context in which this image was created.
        /// </summary>
        internal CalcOrigins CalcOrigin { get; set; }

        internal WebImageSizing? Sizing { get; set; }

        /// <summary>
        /// Address of the cell picture
        /// </summary>
        public ExcelAddress CellAddress { get; internal set; }

        internal bool IsReferenceTo(string wsName, int row, int col)
        {
            return wsName != CellAddress._ws || row != CellAddress._fromRow || col != CellAddress._toCol;
        }

       
    }
}
