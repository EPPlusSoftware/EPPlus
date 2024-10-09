using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.RichValues.LocalImage;
using OfficeOpenXml.RichData.Structures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues
{
    internal static class ExcelRichValueFactory
    {
        public static ExcelRichValue Create(ExcelRichValueStructure structure, int structureId, ExcelRichData richData)
        {
            switch(structure.StructureType)
            {
                case RichDataStructureTypes.ErrorSpill:
                    return new ErrorSpillRichValue(richData);
                case RichDataStructureTypes.ErrorField:
                    return new ErrorFieldRichValue(richData);
                case RichDataStructureTypes.ErrorPropagated:
                    return new ErrorPropagatedRichValue(richData);
                case RichDataStructureTypes.ErrorWithSubType:
                    return new ErrorWithSubTypeRichValue(richData);
                case RichDataStructureTypes.LocalImage:
                    return new LocalImageRichValue(richData);
                case RichDataStructureTypes.LocalImageWithAltText:
                    return new LocalImageAltTextRichValue(richData);
                default:
                    return new ExcelPreserveRichValue(richData, structureId, structure);
            }
        }
    }
}
