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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlus.Fonts.OpenType.FontValidation
{
    public class TableValidationResult
    {
        private readonly List<FontValidationMessage> _messages = new List<FontValidationMessage>();

        public string TableName { get; set; }

        public IEnumerable<FontValidationMessage> Messages => _messages;

        public bool IsValid => !_messages.Any(m => m.Severity == FontValidationSeverity.Error);

        public void AddMessage(FontValidationSeverity severity, string message)
        {
            _messages.Add(new FontValidationMessage(severity, message));
        }

        public IEnumerable<FontValidationMessage> Errors => _messages.Where(m => m.Severity == FontValidationSeverity.Error);
        public IEnumerable<FontValidationMessage> Warnings => _messages.Where(m => m.Severity == FontValidationSeverity.Warning);
        public IEnumerable<FontValidationMessage> Information => _messages.Where(m => m.Severity == FontValidationSeverity.Information);
    }
}
