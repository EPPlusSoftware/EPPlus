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
using EPPlus.Fonts.OpenType.FontValidation;
using System;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Name
{
    internal class NameTableValidator : TableValidatorBase<NameTable>
    {
        public override string TableName
        {
            get { return TableNames.Name; }
        }

        public override Type TableType => typeof(NameTable);

        public override TableValidationResult Validate(NameTable table, FontValidationContext context)
        {
            TableValidationResult result = new TableValidationResult();
            result.TableName = TableName;

            // Rule 1: count vs NameRecords length
            int actualCount = (table.NameRecords == null) ? 0 : table.NameRecords.Length;
            if (table.count != actualCount)
            {
                result.AddMessage(FontValidationSeverity.Error,
                    string.Format("NameTable count ({0}) does not match NameRecords length ({1}).", table.count, actualCount));
            }

            // Rule 2: format check
            if (table.format != 0)
            {
                result.AddMessage(FontValidationSeverity.Warning,
                    string.Format("NameTable format is {0}. Expected 0.", table.format));
            }

            if (table.NameRecords == null || table.NameRecords.Length == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "NameTable has no NameRecords.");
                return result;
            }

            // Find critical NameIDs
            var postScriptNames = table.NameRecords.Where(r => r.nameId == 6).ToList();
            var familyNames = table.NameRecords.Where(r => r.nameId == 1).ToList();
            var subfamilyNames = table.NameRecords.Where(r => r.nameId == 2).ToList();

            // Rule 3: PostScript Name must exist
            if (postScriptNames.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Error, "PostScript Name (NameID 6) is missing.");
            }
            else
            {
                foreach (var psNameRecord in postScriptNames)
                {
                    if (NameRecord.IsNullOrWhiteSpace(psNameRecord.Name))
                    {
                        result.AddMessage(FontValidationSeverity.Error, "PostScript Name is empty or whitespace.");
                    }
                    else if (psNameRecord.Name.Contains(" "))
                    {
                        result.AddMessage(FontValidationSeverity.Error,
                            string.Format("PostScript Name '{0}' contains spaces, which is not allowed.", psNameRecord.Name));
                    }

                    // Check platform for PostScript Name
                    if (psNameRecord.platformId != 3 && psNameRecord.platformId != 0)
                    {
                        result.AddMessage(FontValidationSeverity.Warning,
                            string.Format("PostScript Name '{0}' is not from Windows or Unicode platform (platformId={1}).", psNameRecord.Name, psNameRecord.platformId));
                    }
                }
            }

            // Rule 4: Family and Subfamily should exist
            if (familyNames.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning, "Font Family (NameID 1) is missing.");
            }
            if (subfamilyNames.Count == 0)
            {
                result.AddMessage(FontValidationSeverity.Warning, "Font Subfamily (NameID 2) is missing.");
            }

            // Optional: Check for empty names in critical records
            foreach (var record in familyNames.Concat(subfamilyNames))
            {
                if (NameRecord.IsNullOrWhiteSpace(record.Name))
                {
                    result.AddMessage(FontValidationSeverity.Warning,
                        string.Format("NameID {0} is empty or whitespace.", record.nameId));
                }
            }

            return result;
        }


    }
}
