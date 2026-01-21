/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           TextRun implementation
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Drawing.Text;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// Represents a portion of text with consistent formatting.
    /// </summary>
    public class TextRun
    {
        public string Text { get; set; }
        public MeasurementFont Font { get; set; }
        public int StartIndex { get; set; }
        public int Length { get; set; }
    }
}
