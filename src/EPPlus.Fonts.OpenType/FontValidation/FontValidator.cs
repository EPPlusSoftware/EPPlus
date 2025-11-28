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
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using System.Collections.Generic;


namespace EPPlus.Fonts.OpenType.FontValidation
{
    public class FontValidator
    {
        private readonly List<object> _validators = new List<object>();

        public FontValidator()
        {
            // Register built-in validators
            _validators.Add(new HeadTableValidator());
            _validators.Add(new MaxpTableValidator());
            _validators.Add(new HeadTableValidator());
            _validators.Add(new NameTableValidator());
            _validators.Add(new Os2TableValidator());
        }

        public FontValidationReport Validate(OpenTypeFont font)
        {
            var report = new FontValidationReport();
            var context = new FontValidationContext(font);

            var tableRecordsValidator = new TableRecordsValidator();
            var trResult = tableRecordsValidator.Validate(font, context);
            report.AddResult(trResult);

            foreach (var validator in _validators)
            {
                if (validator is ITableValidator<HeadTable> headValidator)
                {
                    var result = headValidator.Validate(font.HeadTable, context);
                    report.AddResult(result);
                }
                else if (validator is ITableValidator<MaxpTable> maxpValidator)
                {
                    var result = maxpValidator.Validate(font.MaxpTable, context);
                    report.AddResult(result);
                }
                else if (validator is ITableValidator<HheaTable> hheaValidator)
                {
                    var result = hheaValidator.Validate(font.HheaTable, context);
                    report.AddResult(result);
                }
                else if (validator is ITableValidator<NameTable> nameValidator)
                {
                    var result = nameValidator.Validate(font.NameTable, context);
                    report.AddResult(result);
                }
                else if (validator is ITableValidator<Os2Table> os2Validator)
                {
                    var result = os2Validator.Validate(font.Os2Table, context);
                    report.AddResult(result);
                }
            }

            return report;
        }
    }
}
