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
using System.Collections.Generic;
using System.Diagnostics;

namespace EPPlus.Fonts.OpenType.FontValidation
{
    [DebuggerDisplay("IsValid = {IsValid}, Errors = {ErrorCount}, Warnings = {WarningCount}, Tables = {Results.Count}")]
    public class FontValidationReport
    {
        private readonly List<TableValidationResult> _results = new List<TableValidationResult>();

        public IList<TableValidationResult> Results
        {
            get { return _results; }
        }

        public bool IsValid
        {
            get
            {
                foreach (var r in _results)
                {
                    if (!r.IsValid) return false;
                }
                return true;
            }
        }

        // ✅ NYA PROPERTIES

        /// <summary>
        /// Gets all error messages from all tables
        /// </summary>
        public IEnumerable<FontValidationMessage> Errors
        {
            get
            {
                foreach (var result in _results)
                {
                    foreach (var error in result.Errors)
                    {
                        yield return error;
                    }
                }
            }
        }

        /// <summary>
        /// Gets all warning messages from all tables
        /// </summary>
        public IEnumerable<FontValidationMessage> Warnings
        {
            get
            {
                foreach (var result in _results)
                {
                    foreach (var warning in result.Warnings)
                    {
                        yield return warning;
                    }
                }
            }
        }

        /// <summary>
        /// Gets all information messages from all tables
        /// </summary>
        public IEnumerable<FontValidationMessage> Information
        {
            get
            {
                foreach (var result in _results)
                {
                    foreach (var info in result.Information)
                    {
                        yield return info;
                    }
                }
            }
        }

        /// <summary>
        /// Gets total count of errors across all tables
        /// </summary>
        public int ErrorCount
        {
            get
            {
                int count = 0;
                foreach (var result in _results)
                {
                    foreach (var msg in result.Messages)
                    {
                        if (msg.Severity == FontValidationSeverity.Error)
                            count++;
                    }
                }
                return count;
            }
        }

        /// <summary>
        /// Gets total count of warnings across all tables
        /// </summary>
        public int WarningCount
        {
            get
            {
                int count = 0;
                foreach (var result in _results)
                {
                    foreach (var msg in result.Messages)
                    {
                        if (msg.Severity == FontValidationSeverity.Warning)
                            count++;
                    }
                }
                return count;
            }
        }

        public void AddResult(TableValidationResult result)
        {
            _results.Add(result);
        }

        public void AddMessage(FontValidationSeverity severity, string message)
        {
            TableValidationResult globalResult = new TableValidationResult();
            globalResult.TableName = "Font";
            globalResult.AddMessage(severity, message);
            _results.Add(globalResult);
        }

        public string FormatSummary()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Font Validation Summary:");
            foreach (var r in _results)
            {
                sb.AppendLine("[" + r.TableName + "] " + (r.IsValid ? "Valid" : "Invalid"));
                foreach (var msg in r.Messages)
                {
                    sb.AppendLine(msg.ToString());
                }
            }
            return sb.ToString();
        }
    }
}