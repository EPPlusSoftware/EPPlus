/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/19/2026         EPPlus Software AB           Web font subsetting for SVG export
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    /// <summary>
    /// The result of <see cref="WebFontSubsetPlanner.Build"/>: the subsets to declare with @font-face
    /// and the font-family stack for every font that was added.
    /// </summary>
    public sealed class WebFontPlan
    {
        private readonly Dictionary<string, Stack> _stacks;

        internal WebFontPlan(List<WebFontFace> faces, Dictionary<string, Stack> stacks)
        {
            Faces = faces.AsReadOnly();
            _stacks = stacks;
        }

        /// <summary>
        /// The subsets to declare with @font-face, each once, in the order they were first needed.
        /// </summary>
        public IReadOnlyList<WebFontFace> Faces { get; private set; }

        /// <summary>
        /// The family names of the font-family stack for a font added with
        /// <see cref="WebFontSubsetPlanner.AddText"/>, in order and without quotes. Subset names
        /// (<see cref="WebFontFace.FamilyName"/>) and installed family names are mixed; the generic
        /// family is not included, see <see cref="GetGenericFamily"/>. Returns null for a font that
        /// was never added. The family name is matched case-insensitively.
        /// </summary>
        public IReadOnlyList<string> GetFamilyNames(string family, FontSubFamily subFamily)
        {
            Stack stack;
            return _stacks.TryGetValue(WebFontSubsetPlanner.RequestKey(family, subFamily), out stack)
                ? stack.Names
                : null;
        }

        /// <summary>
        /// The family names of the font-family stack to draw one code point of a font added with
        /// <see cref="WebFontSubsetPlanner.AddText"/> with. A code point that measurement routed to a
        /// glyph fallback written by name gets a stack that starts with that fallback, so the browser
        /// does not draw it with another font earlier in the font's stack. Every other code point gets
        /// the font's stack (<see cref="GetFamilyNames(string, FontSubFamily)"/>). Code points that share
        /// a stack get the same list instance, so text can be split where the instance changes.
        /// Returns null for a font that was never added.
        /// </summary>
        public IReadOnlyList<string> GetFamilyNames(string family, FontSubFamily subFamily, int codePoint)
        {
            Stack stack;
            if (!_stacks.TryGetValue(WebFontSubsetPlanner.RequestKey(family, subFamily), out stack))
                return null;

            IReadOnlyList<string> names;
            if (stack.CodePointStacks != null && stack.CodePointStacks.TryGetValue(codePoint, out names))
                return names;
            return stack.Names;
        }

        /// <summary>
        /// The generic family that ends the font-family stack for a font added with
        /// <see cref="WebFontSubsetPlanner.AddText"/>. <see cref="GenericFontFamily.SansSerif"/> for
        /// a font that was never added.
        /// </summary>
        public GenericFontFamily GetGenericFamily(string family, FontSubFamily subFamily)
        {
            Stack stack;
            return _stacks.TryGetValue(WebFontSubsetPlanner.RequestKey(family, subFamily), out stack)
                ? stack.Generic
                : GenericFontFamily.SansSerif;
        }

        internal sealed class Stack
        {
            internal Stack(List<string> names, GenericFontFamily generic, Dictionary<int, IReadOnlyList<string>> codePointStacks)
            {
                Names = names.AsReadOnly();
                Generic = generic;
                CodePointStacks = codePointStacks;
            }

            internal IReadOnlyList<string> Names { get; private set; }
            internal GenericFontFamily Generic { get; private set; }

            /// <summary>
            /// Stacks for code points routed to glyph fallbacks written by name. Null when there are none.
            /// </summary>
            internal Dictionary<int, IReadOnlyList<string>> CodePointStacks { get; private set; }
        }
    }
}