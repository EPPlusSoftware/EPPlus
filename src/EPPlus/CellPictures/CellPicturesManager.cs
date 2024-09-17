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
                var pic = new ExcelCellPicture
                {
                    CellAddress = new ExcelAddress(_sheet.Name, row, col, row, col),
                    ImagePath = relation.Target,
                    CalcOrigin = int.Parse(richData.Values.Last())
                };
                return pic;
            }
            return null;
        }

        public void SetCellPicture(int row, int col, byte[] imageBytes, int calcOrigin = CalcOrigin_AddedByUser)
        {
            var richData = _richDataStore.GetRichData(row, col, ExcelCellPicture.LocalImageStructureType);
            if(richData == null)
            {
                //MetaDataReference mdr;
                //mdr.vm = 
                //_metadataStore.SetValue(row, col)
            }
            using var ms = new MemoryStream(imageBytes);
            ImageInfo imageInfo;
            RichValueRel relation;
            if(_pictureStore.ImageExists(imageBytes))
            {
                imageInfo = _pictureStore.GetImageInfo(imageBytes);
                relation = _richDataStore.GetRelation(imageInfo.Uri.OriginalString, ExcelPackage.schemaImage);
            }
            else
            {
                var pictureType = ImageReader.GetPictureType(ms, true);
                imageInfo = _pictureStore.AddImage(imageBytes, null, pictureType);
            }
            var rdUri = new Uri(ExcelRichValueCollection.PART_URI_PATH, UriKind.Relative);
            var imageUri = UriHelper.GetRelativeUri(rdUri, imageInfo.Uri);
            _richDataStore.AddRichData(ExcelPackage.schemaImage, imageUri.OriginalString, new List<string> { calcOrigin.ToString() }, RichDataStructureFlags.LocalImage, out int vm);
            var md = _sheet._metadataStore.GetValue(row, col);
            md.vm = vm;
            // there should be a #VALUE error in the cell that contains the picture...
            _sheet.Cells[row, col].Value = ExcelErrorValue.Create(eErrorType.Value);
            _sheet._metadataStore.SetValue(row, col, md);

        }

        public void SetCellPicture(int row, int col, Stream imageStream, int calcOrigin = CalcOrigin_AddedByUser)
        {
            imageStream.Seek(0, SeekOrigin.Begin);
            using var sr = new StreamReader(imageStream);
            var bytes = sr.ReadToEnd();
            SetCellPicture(row, col, bytes, calcOrigin);
        }

        public void SetCellPicture(int row, int col, string path, int calcOrigin = CalcOrigin_AddedByUser)
        {
            var imageBytes = File.ReadAllBytes(path);
            SetCellPicture(row, col, imageBytes, calcOrigin);
        }

        public void SetCellPicture(int row, int col, FileInfo fileInfo, int calcOrigin = CalcOrigin_AddedByUser)
        {
            SetCellPicture(row, col, fileInfo.FullName, calcOrigin);
        }

        public void SetCellPicture(int row, int col, ExcelImage image, int calcOrigin = CalcOrigin_AddedByUser)
        {
            SetCellPicture(row, col, image.ImageBytes, calcOrigin);
        }
    }
}
