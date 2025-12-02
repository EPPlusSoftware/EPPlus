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
using System.Diagnostics;

namespace EPPlus.Fonts.OpenType.FontValidation
{
    [DebuggerDisplay("Severity = {Severity}, Messages={Message}")]
    public class FontValidationMessage
    {

        public FontValidationSeverity Severity { get; }
        public string Message { get; }

        public FontValidationMessage(FontValidationSeverity severity, string message)
        {
            Severity = severity;
            Message = message;
        }

    }
}
