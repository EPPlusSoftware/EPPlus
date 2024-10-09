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
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.Structures.Constants
{
    internal static class StructureKeys
    {
        internal static class Errors
        {
            internal static readonly List<ExcelRichValueStructureKey> Propagated =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.PropagatedError.ErrorType, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.PropagatedError.Propagated, RichValueDataType.String)
                ];

            internal static readonly List<ExcelRichValueStructureKey> Field =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.FieldError.ErrorType, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.FieldError.Field, RichValueDataType.String)
                ];

            internal static readonly List<ExcelRichValueStructureKey> Spill =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.Spill.ColOffset, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.Spill.ErrorType, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.Spill.RwOffset, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.Spill.SubType, RichValueDataType.Integer)
                ];

            internal static readonly List<ExcelRichValueStructureKey> WithSubType =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.WithSubType.ErrorType, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.Errors.WithSubType.SubType, RichValueDataType.Integer)
                ];
        }

        internal static class LocalImage
        {
            internal static readonly List<ExcelRichValueStructureKey> Image =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.LocalImages.Image.RelLocalImageIdentifier, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.LocalImages.Image.CalcOrigin, RichValueDataType.Integer)
                ];

            internal static readonly List<ExcelRichValueStructureKey> ImageAltText =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.LocalImages.ImageAltText.RelLocalImageIdentifier, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.LocalImages.ImageAltText.CalcOrigin, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.LocalImages.ImageAltText.Text, RichValueDataType.String)
                ];
        }

        internal static class WebImage
        {
            internal static readonly List<ExcelRichValueStructureKey> Image =
                [
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.WebImageIdentifier, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.Attribution, RichValueDataType.SupportingPropertyBag),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.Text, RichValueDataType.String),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.ComputedImage, RichValueDataType.Bool),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.ImageSizing, RichValueDataType.Integer),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.ImageHeight, RichValueDataType.Decimal),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.ImageWidth, RichValueDataType.Decimal),
                    new ExcelRichValueStructureKey(StructureKeyNames.WebImage.CalcOrigin, RichValueDataType.Integer),
                ];
        }
       
    }
}
