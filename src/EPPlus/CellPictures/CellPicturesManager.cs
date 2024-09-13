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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Metadata;
using OfficeOpenXml.RichData;
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
            _workbook = sheet.Workbook;
            _metadataStore = sheet._metadataStore;
            _metadata = _workbook.Metadata;
            _pictureStore = _workbook._package.PictureStore;
            //var part = _workbook._package.ZipPackage.GetPart(new Uri(PART_URI_PATH, UriKind.Relative));
        }

        private readonly ExcelWorksheet _sheet;
        private readonly ExcelWorkbook _workbook;
        private readonly CellStore<MetaDataReference> _metadataStore;
        private readonly ExcelMetadata _metadata;
        private readonly PictureStore _pictureStore;
        const string LocalImageStructureType = "_localImage";
        const string PART_URI_PATH = "/xl/richData/richValueRel.xml.rels";

        private ExcelRichValue GetRichData(int row, int col)
        {
            var vm = _metadataStore.GetValue(row, col).vm;
            if (vm == 0) return null;
            // vm is a 1-based index pointer
            var vmIx = vm - 1;
            var valueMd = _metadata.ValueMetadata[vmIx];
            var valueRecord = valueMd.Records.First();
            var type = _metadata.MetadataTypes[valueRecord.RecordTypeIndex - 1];
            var futureMetadata = _metadata.MetadataTypes.First(x => x.Name == type.Name);
            return _workbook.RichData.Values.Items[valueRecord.ValueTypeIndex];
        }

        private ExcelMetadataItem CreateMetadataItem()
        {
            //_metadata
            //_workbook.RichData.Values.Items.Add(new ExcelRichValue())
            //var item = new ExcelMetadataItem();
            //item.Records.Add(new ExcelMetadataRecord(_metadata.RichDataTypeIndex, ));
            return null;
        }


        public ExcelCellPicture GetCellPicture(int row, int col)
        {
            var richData = GetRichData(row, col);
            if (richData != null && richData.Structure.Type == LocalImageStructureType)
            {
                var relationIndex = int.Parse(richData.Values.First());
                var relation = _workbook.RichData.GetRelationByIndex(relationIndex);
                var pic = new ExcelCellPicture();
                pic.CellAddress = new ExcelAddress(_sheet.Name, row, col, row, col);
                pic.ImagePath = relation.Target;
                pic.CalcOrigin = int.Parse(richData.Values.Last());
                return pic;
            }
            return null;
        }

        public void SetCellPicture(int row, int col, byte[] imageBytes)
        {
            var richData = GetRichData(row, col);
            if(richData == null)
            {
                //MetaDataReference mdr;
                //mdr.vm = 
                //_metadataStore.SetValue(row, col)
            }
            using var ms = new MemoryStream(imageBytes);
            var pictureType = ImageReader.GetPictureType(ms, true);
            ImageInfo imageInfo;
            if(_pictureStore.ImageExists(imageBytes))
            {
                imageInfo = _pictureStore.GetImageInfo(imageBytes);
            }
            else
            {
                imageInfo = _pictureStore.AddImage(imageBytes, null, pictureType);
            }
        }

        public void SetCellPicture(int row, int col, string path)
        {
            var imageBytes = File.ReadAllBytes(path);
            SetCellPicture(row, col, imageBytes);
        }
    }
}
