using OfficeOpenXml.RichData.IndexRelations;
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues.Errors
{
    internal static class RichValueErrorFactory
    {
        public static ErrorRichValueBase CreateRichValueErrorFromRichData(ExcelRichValue rv, RichDataIndexStore store, RichDataDatabase richDataDb)
        {
            if (rv.Structure == null) return null;
            if (rv.Structure.Type != StructureTypes.Error) return null;
            var flag = RichValueStructureFactory.GetFlag("_error", rv.Structure.Keys);
            if ((flag & RichDataStructureTypes.ErrorSpill) == RichDataStructureTypes.ErrorSpill)
            {
                return new ErrorSpillRichValue(richDataDb, rv.Values);
            }
            else if((flag & RichDataStructureTypes.ErrorPropagated) == RichDataStructureTypes.ErrorPropagated)
            {
                return new ErrorPropagatedRichValue(richDataDb, rv.Values);
            }
            else if((flag & RichDataStructureTypes.ErrorField) == RichDataStructureTypes.ErrorField)
            {
                return new ErrorFieldRichValue(richDataDb, rv.Values);
            }
            else if((flag & RichDataStructureTypes.ErrorWithSubType) == RichDataStructureTypes.ErrorWithSubType)
            {
                return new ErrorWithSubTypeRichValue(richDataDb, rv.Values);
            }
            else return null;
        }
    }
}
