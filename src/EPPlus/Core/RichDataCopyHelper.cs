/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/11/2024         EPPlus Software AB           EPPlus v8
 *************************************************************************************************/
using OfficeOpenXml.CellPictures;
using OfficeOpenXml.Utils;
using OfficeOpenXml.RichData;
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core
{
    internal class RichDataCopyHelper
    {
        public RichDataCopyHelper(ExcelRangeBase sourceRange, ExcelRangeBase destination)
        {
            _sourceRange = sourceRange;
            _destinationRange = destination;
        }

        private ExcelRangeBase _sourceRange;
        private ExcelRangeBase _destinationRange;

        internal void Copy(ExcelRangeCopyOptionFlags? flags = null)
        {
            // do nothing if source and destination is in the same ExcelPackage instance.
            if (_sourceRange._workbook._package == _destinationRange._workbook._package) return;

            var sourcePackage = _sourceRange._workbook._package;
            var sourceSheet = _sourceRange.Worksheet;
            var destPackage = _destinationRange._workbook._package;
            var destSheet = _destinationRange.Worksheet;
            var sourceRichData = _sourceRange._workbook.RichData;

            var maxRow = _sourceRange._toRow > sourceSheet.Dimension?._toRow ? sourceSheet.Dimension._toRow : _sourceRange._toRow;
            var maxCol = _sourceRange._toCol > sourceSheet.Dimension?._toCol ? sourceSheet.Dimension._toCol : _sourceRange._toCol;

            var range = sourceSheet.Cells[_sourceRange._fromRow, _sourceRange._fromCol, maxRow, maxCol];
            foreach(var cell in range)
            {
                var md = sourceSheet._metadataStore.GetValue(cell._fromRow, cell._fromCol);
                if (md.vm > 0)
                {
                    //Copy value metadata of supported rich data types
                    var sourceRv = sourceSheet._richDataStore.GetRichValue(md.vm);
                    sourceRv.SetStructure(sourceRichData.Db);
                    // Local image
                    if (sourceRv.Structure.Type == StructureTypes.LocalImage && (!flags.HasValue || EnumUtil.HasNotFlag(flags.Value, ExcelRangeCopyOptionFlags.ExcludeLocalCellPictures)))
                    {
                        var pic = cell.Picture.Get();
                        var cm = new CellPicturesManager(destSheet);
                        cm.SetCellPicture(cell._fromRow, cell._fromCol, pic.GetImageBytes(), pic.AltText, pic.CalcOrigin);
                    }
                    // Web image
                    else if (sourceRv.Structure.Type == StructureTypes.WebImage && (!flags.HasValue || EnumUtil.HasNotFlag(flags.Value, ExcelRangeCopyOptionFlags.ExcludeWebPictures)))
                    {
                        var pic = cell.Picture.Get();
                        var cm = new CellPicturesManager(destSheet);
                        cm.SetWebPicture(cell._fromRow, cell._fromCol, pic.ExternalAddress, pic.GetImageBytes(), pic.AltText, pic.CalcOrigin);
                    }
                }
                if (md.cm > 0)
                {
                    // copy cell metadata
                }
            }      
        }
    }
}
