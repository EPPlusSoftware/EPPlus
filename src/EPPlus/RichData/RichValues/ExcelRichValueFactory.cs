using OfficeOpenXml.RichData.IndexRelations;
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
        public static ExcelRichValue Create(ExcelRichValueStructure structure, uint structureId, RichDataIndexStore store, ExcelRichData richData)
        {
            switch(structure.StructureType)
            {
                case RichDataStructureTypes.ErrorSpill:
                    return new ErrorSpillRichValue(store, richData);
                case RichDataStructureTypes.ErrorField:
                    return new ErrorFieldRichValue(store, richData);
                case RichDataStructureTypes.ErrorPropagated:
                    return new ErrorPropagatedRichValue(store, richData);
                case RichDataStructureTypes.ErrorWithSubType:
                    return new ErrorWithSubTypeRichValue(store, richData);
                case RichDataStructureTypes.LocalImage:
                    return new LocalImageRichValue(store, richData);
                case RichDataStructureTypes.LocalImageWithAltText:
                    return new LocalImageAltTextRichValue(store, richData);
                default:
                    return new ExcelPreserveRichValue(store, richData, structureId, structure);
            }
        }
    }
}
