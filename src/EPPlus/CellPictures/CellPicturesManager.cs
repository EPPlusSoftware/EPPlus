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
        }

        private readonly ExcelWorksheet _sheet;
        private readonly RichDataStore _richDataStore;
        private readonly PictureStore _pictureStore;
        private static readonly ePictureType[] _validPictureTypes = { ePictureType.Png, ePictureType.Jpg, ePictureType.Gif, ePictureType.Bmp, ePictureType.WebP, ePictureType.Tif, ePictureType.Ico };


        public ExcelCellPicture GetCellPicture(int row, int col)
        {
            var richData = _richDataStore.GetRichValue(row, col, out int? rvIx, StructureTypes.LocalImage);
            if (richData != null)
            {
                var relationIndex = int.Parse(richData.Values.First());
                var relation = _richDataStore.GetRelation(relationIndex);
                var sourceUri = _sheet.Workbook.RichData.RichValueRels.Part.Uri;
                if (richData.Structure.Keys.Count == 2)
                {
                    var pic = new ExcelCellPicture
                    {
                        CellAddress = new ExcelAddress(_sheet.Name, row, col, row, col),
                        ImageUri = UriHelper.ResolvePartUri(sourceUri, relation.TargetUri),
                        CalcOrigin = int.Parse(richData.Values.Last())
                    };
                    return pic;
                }
                else if(richData.Structure.Keys.Count == 3)
                {
                    var pic = new ExcelCellPicture
                    {
                        CellAddress = new ExcelAddress(_sheet.Name, row, col, row, col),
                        ImageUri = UriHelper.ResolvePartUri(sourceUri, relation.TargetUri),
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
            return _validPictureTypes.Any(x => x == type);
        }

        private ExcelRichValue GetImageRichValue(int relIx, CalcOrigins calcOrigin, string altText)
        {
            if (!string.IsNullOrEmpty(altText))
            {
                return new LocalImageAltTextRichValue(_sheet.Workbook)
                {
                    RelLocalImageIdentifier = relIx,
                    CalcOrigin = calcOrigin,
                    Text = altText
                };
            }
            else
            {
                return new LocalImageRichValue(_sheet.Workbook)
                {
                    RelLocalImageIdentifier = relIx,
                    CalcOrigin = calcOrigin
                };
            }
        }

        public void SetCellPicture(int row, int col, byte[] imageBytes, string altText, CalcOrigins calcOrigin = CalcOrigins.StandAlone)
        {
            using var ms = new MemoryStream(imageBytes);
            ImageInfo imageInfo;
            RichValueRel relation = default;
            if(_pictureStore.ImageExists(imageBytes))
            {
                imageInfo = _pictureStore.GetImageInfo(imageBytes);
                relation = _richDataStore.GetRelation(imageInfo.Uri, ExcelPackage.schemaImage);
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
            var richDataValue = _richDataStore.GetRichValue(row, col, out int? rvIndex, StructureTypes.LocalImage);
            var structureType = string.IsNullOrEmpty(altText) ? RichDataStructureTypes.LocalImage : RichDataStructureTypes.LocalImageWithAltText;
            var rdUri = new Uri(ExcelRichValueCollection.PART_URI_PATH, UriKind.Relative);
            var imageUri = UriHelper.GetRelativeUri(rdUri, imageInfo.Uri);
            var md = _sheet._metadataStore.GetValue(row, col);

            var relIx = _richDataStore.CreateRichValueRelation(structureType, imageUri);

            int? valueMetadataIndex = default;
            if (richDataValue == null)
            {
                var imageRichValue = GetImageRichValue(relIx, calcOrigin, altText);
                _richDataStore.AddRichData(imageRichValue, structureType, out int vm);
                valueMetadataIndex = vm;
            }
            else
            {
                var existingPic = GetCellPicture(row, col);
                var imageRichValue = GetImageRichValue(relIx, calcOrigin, altText);
                _richDataStore.UpdateRichData(rvIndex.Value, imageRichValue, imageUri);
                valueMetadataIndex = md.vm;
                if (existingPic != null)
                {
                    _pictureStore.RemoveReference(existingPic.ImageUri);
                }
            }
            md.vm = valueMetadataIndex ?? 0;
            // there should be a #VALUE error in the cell that contains the picture...
            _sheet.Cells[row, col].Value = ExcelErrorValue.Create(eErrorType.Value);
            _sheet._metadataStore.SetValue(row, col, md);
        }

        public void SetCellPicture(int row, int col, Stream imageStream, string altText, CalcOrigins calcOrigin = CalcOrigins.StandAlone)
        {
            imageStream.Seek(0, SeekOrigin.Begin);
            using var sr = new StreamReader(imageStream);
            var bytes = sr.ReadToEnd();
            SetCellPicture(row, col, bytes, altText, calcOrigin);
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

        public void RemoveCellPicture(int row, int col)
        {
            if (!_richDataStore.HasRichData(row, col, out int vm)) return;
            var richData = _richDataStore.GetRichValue(row, col, StructureTypes.LocalImage);
            if(richData != null)
            {
                _richDataStore.DeleteRichData(row, col);
            }
        }
    }
}
