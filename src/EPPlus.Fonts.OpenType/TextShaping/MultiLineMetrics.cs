/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/15/2025         EPPlus Software AB           Initial implementation
 *************************************************************************************************/

namespace EPPlus.Fonts.OpenType.TextShaping
{
    /// <summary>
    /// Metrics for multi-line text measurement
    /// </summary>
    public struct MultiLineMetrics
    {
        /// <summary>
        /// Maximum width of all lines
        /// </summary>
        public float Width { get; set; }

        /// <summary>
        /// Total height (line count × line height)
        /// </summary>
        public float Height { get; set; }

        /// <summary>
        /// Font height without line spacing (ascent + descent)
        /// </summary>
        public float FontHeight { get; set; }

        /// <summary>
        /// Number of lines
        /// </summary>
        public int LineCount { get; set; }

        /// <summary>
        /// Height of a single line (ascent + descent + line gap)
        /// </summary>
        public float LineHeight { get; set; }
    }
}