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
    internal abstract class TableValidatorBase<T> : ITableValidator<T>
        where T : FontTableBase
    {
        public abstract Type TableType { get; }
        public abstract string TableName { get; }

        public FontValidationSeverity LogLevel { get; set; } = FontValidationSeverity.All;

        public abstract TableValidationResult Validate(T table, FontValidationContext context);

        TableValidationResult ITableValidator.Validate(FontTableBase table, FontValidationContext context)
               => Validate((T)table, context);

    }
}
