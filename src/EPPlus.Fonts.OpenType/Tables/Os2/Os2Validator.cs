
using System;
using EPPlus.Fonts.OpenType.FontValidation;

namespace EPPlus.Fonts.OpenType.Tables.Os2
{
    public class Os2TableValidator : ITableValidator<Os2Table>
    {
        public string TableName
        {
            get { return TableNames.Os2; }
        }

        public Type TableType => typeof(Os2Table);

        TableValidationResult ITableValidator.Validate(FontTableBase table, FontValidationContext context)
            => Validate((Os2Table)table, context);


        public TableValidationResult Validate(Os2Table table, FontValidationContext context)
        {
            var result = new TableValidationResult();
            result.TableName = TableName;

            // Rule 1: Version check
            if (table.version < 0 || table.version > 5)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    string.Format("Unexpected OS/2 version: {0}. Expected 0–5.", table.version));
            }

            // Rule 2: usWeightClass (100–900)
            if (table.usWeightClass < 100 || table.usWeightClass > 900)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("usWeightClass out of range: {0}. Expected 100–900.", table.usWeightClass));
            }

            // Rule 3: usWidthClass (1–9)
            if (table.usWidthClass < 1 || table.usWidthClass > 9)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("usWidthClass out of range: {0}. Expected 1–9.", table.usWidthClass));
            }

            // Rule 4: fsType embedding permissions
            if ((table.fsType & 0x0002) != 0)
            {
                result.AddMessage(FontValidationSeverity.Information,
                    "Font has restricted embedding (fsType bit 1 set).");
            }

            // Rule 5: Subscript/Superscript sizes > 0
            if (table.ySubscriptXSize <= 0 || table.ySuperscriptXSize <= 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    "Subscript/Superscript sizes should be greater than zero.");
            }

            // Rule 6: Panose check (warn if all zeros)
            bool panoseAllZero = true;
            foreach (var b in table.panose)
            {
                if (b != 0) { panoseAllZero = false; break; }
            }
            if (panoseAllZero)
            {
                result.AddMessage(FontValidationSeverity.Warning, "Panose classification is all zeros.");
            }


            // Cross-table: usFirstCharIndex/usLastCharIndex vs cmap
            if (context.Font.CmapTable != null)
            {
                int minChar = context.Font.CmapTable.GetMinCharCode();
                int maxChar = context.Font.CmapTable.GetMaxCharCode();
                if (minChar < table.usFirstCharIndex || maxChar > table.usLastCharIndex)
                {
                    result.AddMessage(FontValidationSeverity.Error,
                        string.Format("Cmap character range ({0}-{1}) is outside OS/2 declared range ({2}-{3}).",
                            minChar, maxChar, table.usFirstCharIndex, table.usLastCharIndex));
                }
            }


            return result;
        }
    }
}
