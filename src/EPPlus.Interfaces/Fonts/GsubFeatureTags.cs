/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Public feature selection for GSUB shaping
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Converts <see cref="GsubFeature"/> flag combinations to the OpenType feature tags
    /// <see cref="ShapingOptions.GsubFeatures"/> expects.
    /// </summary>
    public static class GsubFeatureTags
    {
        /// <summary>
        /// Converts the given flags to a list of OpenType feature tags.
        /// </summary>
        /// <remarks>
        /// <see cref="GsubFeature.None"/> currently produces an EMPTY list, not a list that
        /// actively blocks every feature. <see cref="ShapingOptions.GsubFeatures"/> treats a null
        /// or empty list as "apply every feature the font defines" (see its own XML doc), so
        /// passing the result of <c>ToTagList(GsubFeature.None)</c> straight into
        /// <c>ShapingOptions.GsubFeatures</c> does not suppress ligatures - it does the opposite.
        /// Callers that need to guarantee no GSUB substitutions run should set
        /// <c>ShapingOptions.ApplySubstitutions = false</c> instead.
        /// </remarks>
        public static List<string> ToTagList(GsubFeature features)
        {
            var tags = new List<string>();

            if ((features & GsubFeature.Liga) != 0) tags.Add("liga");
            if ((features & GsubFeature.Clig) != 0) tags.Add("clig");
            if ((features & GsubFeature.Dlig) != 0) tags.Add("dlig");

            return tags;
        }
    }
}