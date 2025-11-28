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
using System.Collections.Generic;

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

    public void AddResult(TableValidationResult result)
    {
        _results.Add(result);
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
