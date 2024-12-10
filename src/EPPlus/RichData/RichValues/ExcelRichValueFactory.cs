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
using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.RichValues.Errors;
using OfficeOpenXml.RichData.RichValues.LocalImage;
using OfficeOpenXml.RichData.RichValues.WebImages;
using OfficeOpenXml.RichData.Structures;

namespace OfficeOpenXml.RichData.RichValues
{
    internal static class ExcelRichValueFactory
    {
        public static ExcelRichValue Create(ExcelRichValueStructure structure, uint structureId, RichDataIndexStore store, RichDataDatabase richDataDb)
        {
            switch(structure.StructureType)
            {
                case RichDataStructureTypes.ErrorSpill:
                    return new ErrorSpillRichValue(richDataDb);
                case RichDataStructureTypes.ErrorField:
                    return new ErrorFieldRichValue(richDataDb);
                case RichDataStructureTypes.ErrorPropagated:
                    return new ErrorPropagatedRichValue(richDataDb);
                case RichDataStructureTypes.ErrorWithSubType:
                    return new ErrorWithSubTypeRichValue(richDataDb);
                case RichDataStructureTypes.LocalImage:
                    return new LocalImageRichValue(richDataDb);
                case RichDataStructureTypes.WebImage:
                    return new WebImageRichValue(richDataDb);
                default:
                    return new ExcelPreserveRichValue(richDataDb, structureId, structure);
            }
        }
    }
}
