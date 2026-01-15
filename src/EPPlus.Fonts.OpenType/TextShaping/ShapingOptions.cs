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
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.TextShaping
{
    /// <summary>
    /// Options for controlling text shaping behavior.
    /// </summary>
    public class ShapingOptions
    {
        /// <summary>
        /// Whether to apply GSUB substitutions (ligatures, contextual alternates, etc.).
        /// </summary>
        public bool ApplySubstitutions { get; set; }

        /// <summary>
        /// List of GSUB features to apply (e.g., "liga", "calt", "clig").
        /// If null or empty, all available features will be applied.
        /// </summary>
        public List<string> GsubFeatures { get; set; }

        /// <summary>
        /// Whether to apply GPOS positioning (kerning, mark positioning, etc.).
        /// </summary>
        public bool ApplyPositioning { get; set; }

        /// <summary>
        /// List of GPOS features to apply (e.g., "kern", "mark").
        /// If null or empty, all available features will be applied.
        /// </summary>
        public List<string> GposFeatures { get; set; }

        /// <summary>
        /// Script tag for shaping (e.g., "latn" for Latin).
        /// If null, will use default script.
        /// </summary>
        public string Script { get; set; }

        /// <summary>
        /// Language tag for shaping (e.g., "SWE " for Swedish).
        /// If null, will use default language.
        /// </summary>
        public string Language { get; set; }

        /// <summary>
        /// Default shaping options: ligatures and kerning enabled.
        /// </summary>
        public static ShapingOptions Default
        {
            get
            {
                return new ShapingOptions
                {
                    ApplySubstitutions = true,
                    GsubFeatures = new List<string> { "liga" },
                    ApplyPositioning = true,
                    GposFeatures = new List<string> { "kern" },
                    Script = "latn",
                    Language = null
                };
            }
        }

        /// <summary>
        /// Fast shaping: no substitutions, only kerning.
        /// Use for simple text measurement where ligatures are not important.
        /// </summary>
        public static ShapingOptions Fast
        {
            get
            {
                return new ShapingOptions
                {
                    ApplySubstitutions = false,
                    GsubFeatures = null,
                    ApplyPositioning = true,
                    GposFeatures = new List<string> { "kern" },
                    Script = "latn",
                    Language = null
                };
            }
        }

        /// <summary>
        /// Full shaping: all features enabled.
        /// Use for high-quality rendering.
        /// </summary>
        public static ShapingOptions Full
        {
            get
            {
                return new ShapingOptions
                {
                    ApplySubstitutions = true,
                    GsubFeatures = new List<string> { "liga", "calt", "clig" },
                    ApplyPositioning = true,
                    GposFeatures = new List<string> { "kern", "mark" },
                    Script = "latn",
                    Language = null
                };
            }
        }

        /// <summary>
        /// No shaping: just map characters to glyphs.
        /// Fastest option, but no ligatures or kerning.
        /// </summary>
        public static ShapingOptions None
        {
            get
            {
                return new ShapingOptions
                {
                    ApplySubstitutions = false,
                    GsubFeatures = null,
                    ApplyPositioning = false,
                    GposFeatures = null,
                    Script = null,
                    Language = null
                };
            }
        }

        public ShapingOptions()
        {
            ApplySubstitutions = true;
            ApplyPositioning = true;
            GsubFeatures = new List<string>();
            GposFeatures = new List<string>();
        }
    }
}