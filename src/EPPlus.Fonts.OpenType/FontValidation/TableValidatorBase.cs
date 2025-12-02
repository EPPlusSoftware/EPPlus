using EPPlus.Fonts.OpenType.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
