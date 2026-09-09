/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Public feature selection for GPOS shaping
 *************************************************************************************************/
using System.Collections.Generic;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Converts <see cref="GposFeature"/> flag combinations to the OpenType feature tags
    /// <see cref="ShapingOptions.GposFeatures"/> expects.
    /// </summary>
    public static class GposFeatureTags
    {
        /// <summary>
        /// Converts the given flags to a list of OpenType feature tags.
        /// </summary>
        /// <remarks>
        /// <see cref="GposFeature.None"/> currently produces an EMPTY list, not a list that
        /// actively blocks every feature. <see cref="ShapingOptions.GposFeatures"/> treats a null
        /// or empty list as "apply every feature the font defines" (see
        /// <c>TextShaper.ApplyPositioning</c>'s <c>applyAllFeatures</c> check), so passing the
        /// result of <c>ToTagList(GposFeature.None)</c> straight into
        /// <c>ShapingOptions.GposFeatures</c> does not suppress kerning/mark positioning - it does
        /// the opposite. Callers that need to guarantee no GPOS positioning runs should set
        /// <c>ShapingOptions.ApplyPositioning = false</c> instead.
        /// </remarks>
        public static List<string> ToTagList(GposFeature features)
        {
            var tags = new List<string>();

            if ((features & GposFeature.Kern) != 0) tags.Add("kern");
            if ((features & GposFeature.Mark) != 0) tags.Add("mark");

            return tags;
        }
    }
}