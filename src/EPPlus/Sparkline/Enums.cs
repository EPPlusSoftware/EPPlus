/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Sparkline
{
    /// <summary>
    /// Show empty cells as
    /// </summary>
    public enum eDispBlanksAs
    {
        /// <summary>
        /// Connect datapoints with line
        /// </summary>
        Span,
        /// <summary>
        /// A gap
        /// </summary>
        Gap,
        /// <summary>
        /// As Zero
        /// </summary>
        Zero
    }
    /// <summary>
    /// Type of sparkline
    /// </summary>
    public enum eSparklineType
    {
        /// <summary>
        /// Line Sparkline
        /// </summary>
        Line,
        /// <summary>
        /// Column Sparkline
        /// </summary>
        Column,
        /// <summary>
        /// Win/Loss Sparkline
        /// </summary>
        Stacked
    }
    /// <summary>
    /// Axis min/max settings
    /// </summary>
    public enum eSparklineAxisMinMax
    {
        /// <summary>
        /// Individual per sparklines
        /// </summary>
        Individual,
        /// <summary>
        /// Same for all sparklines
        /// </summary>
        Group,
        /// <summary>
        /// A custom value
        /// </summary>
        Custom
    }
}
