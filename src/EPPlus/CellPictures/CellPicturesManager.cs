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
using OfficeOpenXml.Constants;
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.RichData;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using static OfficeOpenXml.ExcelWorksheet;

namespace OfficeOpenXml.CellPictures
{
    internal class CellPicturesManager
    {
        public CellPicturesManager(ExcelWorksheet sheet)
        {
            _sheet = sheet;
            _richDataStore = new RichDataStore(sheet);
            _pictureStore = sheet.Workbook._package.PictureStore;
        }

        private readonly ExcelWorksheet _sheet;
        private readonly RichDataStore _richDataStore;
        private readonly PictureStore _pictureStore;
        // CalcOrigin values:
        const int CalcOrigin_ByReference = 4;
        const int CalcOrigin_AddedByUser = 5;


        public ExcelCellPicture GetCellPicture(int row, int col)
        {
            var richData = _richDataStore.GetRichData(row, col, ExcelCellPicture.LocalImageStructureType);
            if (richData != null)
            {
                var relationIndex = int.Parse(richData.Values.First());
                var relation = _richDataStore.GetRelation(relationIndex);
                if(richData.Structure.Keys.Count == 2)
                {
                    var pic = new ExcelCellPicture
                    {
                        CellAddress = new ExcelAddress(_sheet.Name, row, col, row, col),
                        ImagePath = relation.Target,
                        CalcOrigin = int.Parse(richData.Values.Last())
                    };
                    return pic;
                }
                else if(richData.Structure.Keys.Count == 3)
                {
                    var pic = new ExcelCellPicture
                    {
                        CellAddress = new ExcelAddress(_sheet.Name, row, col, row, col),
                        ImagePath = relation.Target,
                        CalcOrigin = int.Parse(richData.Values[1]),
                        AltText = richData.Values.Last()
                    };
                    return pic;
                }
               
            }
            return null;
        }

        private bool IsValidPictureType(ePictureType type)
        {
            return type == ePictureType.Png || type == ePictureType.Jpg || type == ePictureType.Gif || type == ePictureType.Bmp || type == ePictureType.WebP || type == ePictureType.Tif || type == ePictureType.Ico;
        }

        public void SetCellPicture(int row, int col, byte[] imageBytes, string altText, CalcOrigins calcOrigin = CalcOrigins.AddedByUserAltText)
        {
            using var ms = new MemoryStream(imageBytes);
            ImageInfo imageInfo;
            RichValueRel relation = default;
            if(_pictureStore.ImageExists(imageBytes))
            {
                imageInfo = _pictureStore.GetImageInfo(imageBytes);
                relation = _richDataStore.GetRelation(imageInfo.Uri.OriginalString, ExcelPackage.schemaImage);
            }
            else
            {
                var pictureType = ImageReader.GetPictureType(ms, true);
                if(pictureType == null)
                {
                    throw new ArgumentException("Image type not supported/identified.");
                }
                else if (!IsValidPictureType(pictureType.Value))
                {
                    throw new ArgumentException($"'{pictureType.Value}' is not a supported image type for in-cell pictures. Use png, jpg, gif, bmp, webp, tif or ico.");
                }
                imageInfo = _pictureStore.AddImage(imageBytes, null, pictureType);
            }
            var richDataValue = _richDataStore.GetRichData(row, col, ExcelCellPicture.LocalImageStructureType);
            var flag = string.IsNullOrEmpty(altText) ? RichDataStructureFlags.LocalImage : RichDataStructureFlags.LocalImageWithAltText;
            var rdUri = new Uri(ExcelRichValueCollection.PART_URI_PATH, UriKind.Relative);
            var imageUri = UriHelper.GetRelativeUri(rdUri, imageInfo.Uri);
            int valueMetadataIndex = -1;
            if(richDataValue == null)
            {
                _richDataStore.AddRichData(ExcelPackage.schemaImage, imageUri.OriginalString, new List<string> { ((int)calcOrigin).ToString(), altText }, flag, out int vm);
                valueMetadataIndex = 1;
            }
            else
            {
                _richDataStore.UpdateRichData(richDataValue, ExcelPackage.schemaImage, imageUri.OriginalString, new List<string> { ((int)calcOrigin).ToString(), altText }, flag);
            }
            
            
            var md = _sheet._metadataStore.GetValue(row, col);
            md.vm = valueMetadataIndex;
            // there should be a #VALUE error in the cell that contains the picture...
            _sheet.Cells[row, col].Value = ExcelErrorValue.Create(eErrorType.Value);
            _sheet._metadataStore.SetValue(row, col, md);

        }

        public void SetCellPicture(int row, int col, Stream imageStream, string altText, CalcOrigins calcOrigin = CalcOrigins.AddedByUserAltText)
        {
            imageStream.Seek(0, SeekOrigin.Begin);
            using var sr = new StreamReader(imageStream);
            var bytes = sr.ReadToEnd();
            SetCellPicture(row, col, bytes, altText, calcOrigin);
        }

        public void SetCellPicture(int row, int col, string path, string altText, CalcOrigins calcOrigin = CalcOrigins.AddedByUserAltText)
        {
            var imageBytes = File.ReadAllBytes(path);
            SetCellPicture(row, col, imageBytes, altText, calcOrigin);
        }

        public void SetCellPicture(int row, int col, FileInfo fileInfo, string altText, CalcOrigins calcOrigin = CalcOrigins.AddedByUserAltText)
        {
            SetCellPicture(row, col, fileInfo.FullName, altText, calcOrigin);
        }

        public void SetCellPicture(int row, int col, ExcelImage image, string altText, CalcOrigins calcOrigin = CalcOrigins.AddedByUserAltText)
        {
            SetCellPicture(row, col, image.ImageBytes, altText, calcOrigin);
        }
    }
}
