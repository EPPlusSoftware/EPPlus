﻿/*************************************************************************************************
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
using OfficeOpenXml.RichData.Structures;
using OfficeOpenXml.RichData.Structures.Constants;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.RichData.RichValues.Errors
{
    internal class ErrorFieldRichValue : ExcelRichValue
    {
        public ErrorFieldRichValue(RichDataIndexStore store, ExcelRichData richData) : base(store, richData, RichDataStructureTypes.ErrorField)
        {
        }

        public string Field
        {
            get
            {
                return GetValue(StructureKeyNames.Errors.FieldError.Field);
            }
            set
            {
                SetValue(StructureKeyNames.Errors.FieldError.Field, value);
            }
        }
    }
}
