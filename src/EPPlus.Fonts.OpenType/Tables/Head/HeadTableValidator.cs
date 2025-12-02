
using System;
using EPPlus.Fonts.OpenType.FontValidation;

namespace EPPlus.Fonts.OpenType.Tables.Head
{
    internal class HeadTableValidator : TableValidatorBase<HeadTable>
    {
        public override string TableName
        {
            get { return TableNames.Head; }
        }

        public override Type TableType => typeof(HeadTable);


        public override TableValidationResult Validate(HeadTable table, FontValidationContext context)
        {
            var result = new TableValidationResult();
            result.TableName = TableName;

            // Magic number
            if (table.MagicNumber != HeadTableConstants.MagicNumber)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format($"Invalid magic number: {0:X}. Expected {HeadTableConstants.MagicNumber}.", table.MagicNumber));
            }

            // unitsPerEm range
            if (table.UnitsPerEm < HeadTableConstants.UnitsPerEmMin || table.UnitsPerEm > HeadTableConstants.UnitsPerEmMax)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    $"unitsPerEm out of range: {table.UnitsPerEm}. Expected {HeadTableConstants.UnitsPerEmMin}]–{HeadTableConstants.UnitsPerEmMax}.");
            }

            // indexToLocFormat
            if (table.IndexToLocFormat != HeadTable.IndexToLocFormats.Offset16 &&
                table.IndexToLocFormat != HeadTable.IndexToLocFormats.Offset32)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("Invalid indexToLocFormat: {0}. Expected Offset16 or Offset32.", table.IndexToLocFormat));
            }

            // Bounding box
            if (table.Xmax < table.Xmin)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("xMax ({0}) is less than xMin ({1}).", table.Xmax, table.Xmin));
            }
            if (table.Ymax < table.Ymin)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("yMax ({0}) is less than yMin ({1}).", table.Ymax, table.Ymin));
            }

            // Created/Modified dates
            if (table.Created < LongDateTimeLimits.MinSeconds || table.Created > LongDateTimeLimits.MaxSeconds)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    string.Format("Created date seems invalid: {0}.", table.Created));
            }
            if (table.Modified < LongDateTimeLimits.MinSeconds || table.Modified > LongDateTimeLimits.MaxSeconds)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    string.Format("Modified date seems invalid: {0}.", table.Modified));
            }

            if(table.Modified < table.Created)
            {
                result.AddMessage(FontValidationSeverity.Information,
                    "Modified LONGDATETIME is earlier than Created LONGDATETIME.");
            }

            // Version info
            if (table.MajorVersion != 1)
            {
                result.AddMessage(FontValidationSeverity.Information,
                    string.Format("Unexpected major version: {0}. Expected 1.", table.MajorVersion));
            }

            return result;
        }
    }

}
