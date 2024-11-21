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
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.RichValues;
using OfficeOpenXml.RichData.RichValues.LocalImage;
using OfficeOpenXml.RichData.RichValues.Relations;
using OfficeOpenXml.RichData.Structures.Constants;
using OfficeOpenXml.RichData.Structures.LocalImages;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
            _referenceCache = sheet.Workbook.CellPictureReferenceCache;
        }

        private readonly ExcelWorksheet _sheet;
        private readonly RichDataStore _richDataStore;
        private readonly PictureStore _pictureStore;
        private readonly CellPictureReferenceCache _referenceCache;
        private static readonly ePictureType[] _validPictureTypes = { ePictureType.Png, ePictureType.Jpg, ePictureType.Gif, ePictureType.Bmp, ePictureType.WebP, ePictureType.Tif, ePictureType.Ico };

        private ExcelCellPicture GetExcelCellPictureByRichValue(ExcelRichValue richValue, int row, int col, uint vmId)
        {
            if (richValue.StructureType == RichDataStructureTypes.LocalImage)
            {
                var rdLi = richValue.As.LocalImage;
                var pic = new ExcelCellPicture(vmId, rdLi.ImageUri, _pictureStore, ExcelCellPictureTypes.LocalImage)
                {
                    CellAddress = new ExcelAddress(_sheet.Name, row, col, row, col),
                    AltText = rdLi.Text,
                    CalcOrigin = rdLi.CalcOrigin ?? CalcOrigins.None
                };
                return pic;
            }
            //else if (richValue.Structure.StructureType == RichDataStructureTypes.LocalImageWithAltText)
            //{
            //    var rdLia = richValue.As.LocalImageAltText;
            //    var pic = new ExcelCellPicture(vmId, rdLia.ImageUri, _pictureStore, ExcelCellPictureTypes.LocalImage)
            //    {
            //        CellAddress = new ExcelAddress(_sheet.Name, row, col, row, col),
            //        CalcOrigin = rdLia.CalcOrigin ?? CalcOrigins.None,
            //        AltText = rdLia.Text
            //    };
            //    return pic;
            //}
            else if (richValue.Structure.StructureType == RichDataStructureTypes.WebImage)
            {
                var rdWi = richValue.As.WebImage;
                var pic = new ExcelCellPicture(vmId, rdWi.ImageUri, _pictureStore, ExcelCellPictureTypes.WebImage)
                {
                    CellAddress = new ExcelAddress(_sheet.Name, row, col, row, col),
                    AltText = rdWi.Text,
                    CalcOrigin = rdWi.CalcOrigin ?? CalcOrigins.None
                };
                return pic;
            }
            return null;
        }

        public ExcelCellPicture GetCellPicture(int row, int col)
        {
            var richValue = _richDataStore.GetRichValue(row, col, out uint vmId, StructureTypes.LocalImage);
            if (richValue != null)
            {
                return GetExcelCellPictureByRichValue(richValue, row, col, vmId);
               
            }
            return null;
        }

        private bool IsValidPictureType(ePictureType type)
        {
            return _validPictureTypes.Any(x => x == type);
        }

        private ExcelRichValue CreateImageRichValue(Uri imageUri, CalcOrigins calcOrigin, string altText)
        {
            //if (!string.IsNullOrEmpty(altText))
            //{
            //    return new LocalImageAltTextRichValue(_sheet.Workbook)
            //    {
            //        ImageUri = imageUri,
            //        CalcOrigin = calcOrigin,
            //        Text = altText
            //    };
            //}
            //else
            //{
            //    return new LocalImageRichValue(_sheet.Workbook)
            //    {
            //        ImageUri = imageUri,
            //        CalcOrigin = calcOrigin
            //    };
            //}
            return new LocalImageRichValue(_sheet.Workbook)
            {
                ImageUri = imageUri,
                CalcOrigin = calcOrigin,
                Text = altText
            };
        }

        public void SetCellPicture(int row, int col, Stream imageStream, string altText, CalcOrigins calcOrigin = CalcOrigins.StandAlone)
        {
            var imageBytes = StreamUtil.ReadStreamToByteArray(imageStream);
            SetCellPicture(row, col, imageBytes, altText, calcOrigin);
        }

        public void SetCellPicture(int row, int col, string path, string altText, CalcOrigins calcOrigin = CalcOrigins.StandAlone)
        {
            var imageBytes = File.ReadAllBytes(path);
            SetCellPicture(row, col, imageBytes, altText, calcOrigin);
        }

        public void SetCellPicture(int row, int col, FileInfo fileInfo, string altText, CalcOrigins calcOrigin = CalcOrigins.StandAlone)
        {
            SetCellPicture(row, col, fileInfo.FullName, altText, calcOrigin);
        }

        public void SetCellPicture(int row, int col, ExcelImage image, string altText, CalcOrigins calcOrigin = CalcOrigins.StandAlone)
        {
            SetCellPicture(row, col, image.ImageBytes, altText, calcOrigin);
        }

        public void SetCellPicture(int row, int col, byte[] imageBytes, string altText, CalcOrigins calcOrigin = CalcOrigins.StandAlone)
        {
            // Add image to picture store and create relation
            var imageInfo = AddToPictureStore(imageBytes);
            
            var rdUri = new Uri(ExcelRichValueCollection.PART_URI_PATH, UriKind.Relative);
            var imageUri = UriHelper.ResolvePartUri(rdUri, imageInfo.Uri);

            if(_referenceCache.Contains(imageUri, calcOrigin, out uint cachedVmId))
            {
                AddReferenceToPicture(row, col, calcOrigin, imageUri, cachedVmId);
                return;
            }

            var hasRv = _richDataStore.HasRichData(row, col, out MetaDataReference md);
            // no existing rich data, add new
            if (!hasRv)
            {
                AddNewPicture(row, col, altText, calcOrigin, imageUri);
            }
            else
            {
                var existingPic = GetCellPicture(row, col);
                if(existingPic != null)
                {
                    if(existingPic.ImageUri.OriginalString == imageUri.OriginalString)
                    {
                        return;
                    }
                    else
                    {
                        _referenceCache.RemoveReference(existingPic.ImageUri, calcOrigin, out int numberOfRefsLeft);
                    }
                }
                else
                {
                    // there was rich data connected to the cell, we leave it as it is
                }
                AddNewPicture(row, col, altText, calcOrigin, imageUri);
            }
        }

        private void AddNewPicture(int row, int col, string altText, CalcOrigins calcOrigin, Uri imageUri)
        {
            var imageRichValue = CreateImageRichValue(imageUri, calcOrigin, altText);
            imageRichValue.SetStructure(_sheet.Workbook.RichData);
            _richDataStore.AddRichData(row, col, imageRichValue, out uint vmId);
            var newPic = GetExcelCellPictureByRichValue(imageRichValue, row, col, vmId);
            SetCellValue(row, col, newPic);
            SetValueMetadata(row, col, vmId);
            if(!_referenceCache.Contains(imageUri, calcOrigin))
            {
                _referenceCache.Add(imageUri, calcOrigin, vmId);
            }
            else
            {
                _referenceCache.AddReference(imageUri, calcOrigin);
            }
        }

        private void AddReferenceToPicture(int row, int col, CalcOrigins calcOrigin, Uri imageUri, uint vmId)
        {
            var rv = _richDataStore.GetRichValue(vmId);
            var pic = GetExcelCellPictureByRichValue(rv, row, col, vmId);
            SetCellValue(row, col, pic);
            _referenceCache.AddReference(imageUri, calcOrigin);
            SetValueMetadata(row, col, vmId);
        }

        private void SetValueMetadata(int row, int col, uint vmId)
        {
            var metadataVals = _sheet._metadataStore.GetValue(row, col);
            metadataVals.vm = vmId;
            _sheet._metadataStore.SetValue(row, col, metadataVals);
        }

        private void SetCellValue(int row, int col, ExcelCellPicture pic)
        {
            if(pic.CalcOrigin == CalcOrigins.Reference)
            {
                _sheet.SetValueInner(row, col, pic);
            }
            else
            {
                _sheet.Cells[row, col].Value = pic;
            }
        }

        /// <summary>
        /// Deletes a cell picture.
        /// </summary>
        /// <param name="row">Cell row</param>
        /// <param name="col">Cell column</param>
        public void DeleteCellPicture(int row, int col)
        {
            _richDataStore.DeleteRichData(row, col);
        }

        private ImageInfo AddToPictureStore(byte[] imageBytes)
        {
            ImageInfo imageInfo;
            if (_pictureStore.ImageExists(imageBytes))
            {
                imageInfo = _pictureStore.GetImageInfo(imageBytes);
            }
            else
            {
                using var ms = new MemoryStream(imageBytes);
                var pictureType = ImageReader.GetPictureType(ms, true);
                if (pictureType == null)
                {
                    throw new ArgumentException("Image type not supported/identified.");
                }
                else if (!IsValidPictureType(pictureType.Value))
                {
                    throw new ArgumentException($"'{pictureType.Value}' is not a supported image type for in-cell pictures. Use png, jpg, gif, bmp, webp, tif or ico.");
                }
                imageInfo = _pictureStore.AddImage(imageBytes, null, pictureType);
            }

            return imageInfo;
        }

        public void RemoveCellPicture(int row, int col)
        {
            if (!_richDataStore.HasRichData(row, col, out uint vmId)) return;
            var pic = _sheet.Cells[row, col].Value as ExcelCellPicture;
            if(pic != null)
            {
                _referenceCache.RemoveReference(pic.ImageUri, pic.CalcOrigin, out int numberOfReferencesLeft);
                if (numberOfReferencesLeft <= 0)
                {
                    _richDataStore.DeleteRichData(row, col);
                }
                var mdr = _sheet._metadataStore.GetValue(row, col);
                mdr.vm = 0;
                _sheet._metadataStore.SetValue(row, col, mdr);
                _sheet.Cells[row, col].Value = null;
            }
        }
    }
}
