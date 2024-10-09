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

namespace OfficeOpenXml.RichData.Structures.SpecialKeysAndFlags
{
    internal static class SpecialKeysItems
    {
        private static readonly Dictionary<string, ExcelRichValueStructureKey> _keys = new Dictionary<string, ExcelRichValueStructureKey>
        {
             { SpecialKeyNames.Attribution, new ExcelRichValueStructureKey(SpecialKeyNames.Attribution, RichValueDataType.SupportingPropertyBag) },
             { SpecialKeyNames.CanonicalPropertyNames, new ExcelRichValueStructureKey(SpecialKeyNames.CanonicalPropertyNames, RichValueDataType.SupportingPropertyBag) },
             { SpecialKeyNames.ClassificationId, new ExcelRichValueStructureKey(SpecialKeyNames.ClassificationId, RichValueDataType.String) },
             { SpecialKeyNames.CRID, new ExcelRichValueStructureKey(SpecialKeyNames.CRID, RichValueDataType.Preserve) },
             { SpecialKeyNames.Display, new ExcelRichValueStructureKey(SpecialKeyNames.Display, RichValueDataType.SupportingPropertyBag) },
             { SpecialKeyNames.Flags, new ExcelRichValueStructureKey(SpecialKeyNames.Flags, RichValueDataType.SupportingPropertyBag) },
             { SpecialKeyNames.Format, new ExcelRichValueStructureKey(SpecialKeyNames.Format, RichValueDataType.SupportingPropertyBag) },
             { SpecialKeyNames.Icon, new ExcelRichValueStructureKey(SpecialKeyNames.Icon, RichValueDataType.String) },
             { SpecialKeyNames.Provider, new ExcelRichValueStructureKey(SpecialKeyNames.Provider, RichValueDataType.SupportingPropertyBag) },
             { SpecialKeyNames.Self, new ExcelRichValueStructureKey(SpecialKeyNames.Self, RichValueDataType.Remove) },
             { SpecialKeyNames.SubLabel, new ExcelRichValueStructureKey(SpecialKeyNames.SubLabel, RichValueDataType.SupportingPropertyBag) },
             { SpecialKeyNames.ViewInfo, new ExcelRichValueStructureKey(SpecialKeyNames.ViewInfo, RichValueDataType.SupportingPropertyBag) },
             { SpecialKeyNames.Order, new ExcelRichValueStructureKey(SpecialKeyNames.Order, RichValueDataType.SupportingPropertyBagArray) },
        };
        public static bool Exists(string key, out ExcelRichValueStructureKey structureKey)
        {
            if(_keys.ContainsKey(key))
            {
                structureKey = _keys[key];
                return true;
            }
            structureKey = null;
            return false;
        }
    }
}
