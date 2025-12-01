/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables;
using System;

namespace EPPlus.Fonts.OpenType.FontValidation
{
    public interface ITableValidator
    {
        // Target table type
        Type TableType { get; }

        // Human-readable table name
        string TableName { get; }

        // Non-generic validate for dispatcher
        TableValidationResult Validate(FontTableBase table, FontValidationContext context);
    }

    // Generic version for validator implementations
    public interface ITableValidator<T> : ITableValidator where T : FontTableBase
    {
        TableValidationResult Validate(T table, FontValidationContext context);
    }

}
