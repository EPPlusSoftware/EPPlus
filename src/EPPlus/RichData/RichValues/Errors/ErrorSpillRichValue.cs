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
using OfficeOpenXml.RichData.Structures.Constants;

namespace OfficeOpenXml.RichData.RichValues.Errors
{
    internal class ErrorSpillRichValue : ErrorRichValueBase
    {

        public ErrorSpillRichValue(RichDataDatabase richDataDb) : base(richDataDb, RichDataStructureTypes.Error | RichDataStructureTypes.ErrorSpill)
        {
        }

        public ErrorSpillRichValue(RichDataDatabase richDataDb, IndexedSubsetCollection<ExcelRichValueValue> values)
            : base(richDataDb, values, RichDataStructureTypes.Error | RichDataStructureTypes.ErrorSpill)
        {
            
        }

        public int? ColOffset
        {
            get
            {
                return GetValueInt(StructureKeyNames.Errors.Spill.ColOffset);
            }
            set
            {
                SetValue(StructureKeyNames.Errors.Spill.ColOffset, value);
            }
        }

        public int? RwOffset
        {
            get
            {
                return GetValueInt(StructureKeyNames.Errors.Spill.RwOffset);
            }
            set
            {
                SetValue(StructureKeyNames.Errors.Spill.RwOffset, value);
            }
        }

        public int? SubType
        {
            get
            {
                return GetValueInt(StructureKeyNames.Errors.Spill.SubType);
            }
            set
            {
                SetValue(StructureKeyNames.Errors.Spill.SubType, value);
            }
        }

        public bool AreEqual(int errorType, int colOffset, int rwOffset, int? subType = default)
        {
            if (subType.HasValue)
            {
                return errorType == ErrorType && colOffset == ColOffset && rwOffset == RwOffset && subType.Value == (SubType ?? int.MinValue);
            }
            return errorType == ErrorType && colOffset == ColOffset && rwOffset == RwOffset;
        }
    }
}
