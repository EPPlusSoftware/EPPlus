using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.RichValues.LocalImage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues
{
    internal static class ExcelRichValueFactory
    {
        public static ExcelRichValue Create(RichDataStructureTypes type, ExcelWorkbook workbook)
        {
            switch(type)
            {
                case RichDataStructureTypes.ErrorSpill:
                    return new ErrorSpillRichValue(workbook);
                case RichDataStructureTypes.ErrorField:
                    return new ErrorFieldRichValue(workbook);
                case RichDataStructureTypes.ErrorPropagated:
                    return new ErrorPropagatedRichValue(workbook);
                case RichDataStructureTypes.ErrorWithSubType:
                    return new ErrorWithSubTypeRichValue(workbook);
                case RichDataStructureTypes.LocalImage:
                    return new LocalImageRichValue(workbook);
                case RichDataStructureTypes.LocalImageWithAltText:
                    return new LocalImageAltTextRichValue(workbook);
                default:
                    return new ExcelPreserveRichValue(workbook);
            }
        }
    }
}
