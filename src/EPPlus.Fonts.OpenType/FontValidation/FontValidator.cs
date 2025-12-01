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
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Hmtx;
using EPPlus.Fonts.OpenType.Tables.Loca;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Name;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Tables.Post;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.FontValidation
{
    public class FontValidator
    {
        private readonly List<ITableValidator> _validators;
        private readonly Dictionary<Type, Func<OpenTypeFont, FontTableBase>> _tableAccessors;

        public FontValidator()
        {
            _validators = new List<ITableValidator>();
            _tableAccessors = new Dictionary<Type, Func<OpenTypeFont, FontTableBase>>();

            // Register validators
            _validators.Add(new HeadTableValidator());
            _validators.Add(new MaxpTableValidator());
            _validators.Add(new HheaTableValidator());
            _validators.Add(new NameTableValidator());
            _validators.Add(new Os2TableValidator());
            _validators.Add(new PostTableValidator());
            _validators.Add(new CmapTableValidator());
            _validators.Add(new LocaTableValidator());
            _validators.Add(new HmtxTableValidator());
            _validators.Add(new GlyfTableValidator());
            // Add more as needed...

            // Register table accessors
            _tableAccessors.Add(typeof(HeadTable), delegate(OpenTypeFont font) { return font.HeadTable; });
            _tableAccessors.Add(typeof(MaxpTable), delegate(OpenTypeFont font) { return font.MaxpTable; });
            _tableAccessors.Add(typeof(HheaTable), delegate(OpenTypeFont font) { return font.HheaTable; });
            _tableAccessors.Add(typeof(NameTable), delegate(OpenTypeFont font) { return font.NameTable; });
            _tableAccessors.Add(typeof(Os2Table), delegate(OpenTypeFont font) { return font.Os2Table; });
            _tableAccessors.Add(typeof(PostTable), delegate (OpenTypeFont font) { return font.PostTable; });
            _tableAccessors.Add(typeof(CmapTable), delegate (OpenTypeFont font) { return font.CmapTable; });
            _tableAccessors.Add(typeof(LocaTable), delegate (OpenTypeFont font) { return font.LocaTable; });
            _tableAccessors.Add(typeof(HmtxTable), delegate (OpenTypeFont font) { return font.HmtxTable; });
            _tableAccessors.Add(typeof(GlyfTable), delegate (OpenTypeFont font) { return font.GlyfTable; });
        }

        public FontValidationReport Validate(OpenTypeFont font)
        {
            FontValidationReport report = new FontValidationReport();
            FontValidationContext context = new FontValidationContext(font);

            // Validate table records first
            TableRecordsValidator tableRecordsValidator = new TableRecordsValidator();
            TableValidationResult trResult = tableRecordsValidator.Validate(font, context);
            report.AddResult(trResult);

            // Validate each registered table
            for (int i = 0; i < _validators.Count; i++)
            {
                ITableValidator validator = _validators[i];
                Func<OpenTypeFont, FontTableBase> accessor;

                if (!_tableAccessors.TryGetValue(validator.TableType, out accessor))
                {
                    report.AddMessage(FontValidationSeverity.Warning,
                        "No accessor registered for table " + validator.TableName + ".");
                    continue;
                }

                FontTableBase table = accessor(font);

                if (table == null)
                {
                    FontValidationSeverity severity = validator.TableType != null && table == null
                        ? FontValidationSeverity.Warning
                        : FontValidationSeverity.Error;

                    // Use IsEssentialTable property for severity
                    if (validator.TableType != null && table == null)
                    {
                        severity = FontValidationSeverity.Warning;
                    }

                    report.AddMessage(table != null && table.IsEssentialTable
                        ? FontValidationSeverity.Error
                        : FontValidationSeverity.Warning,
                        "Table " + validator.TableName + " is missing.");
                    continue;
                }

                // Validate table
                TableValidationResult result = validator.Validate(table, context);
                if (result != null)
                {
                    report.AddResult(result);
                }
                else
                {
                    report.AddMessage(FontValidationSeverity.Error,
                        "Validator for " + validator.TableName + " returned null result.");
                }
            }

            return report;
        }
    }
}
