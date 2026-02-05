/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/26/2026         EPPlus Software AB           StringBuilder extensions for readability
 *************************************************************************************************/
using System.Collections.Generic;
using System.Text;

namespace EPPlus.Fonts.OpenType.Utilities
{
    /// <summary>
    /// Extension methods for StringBuilder to improve code readability in text wrapping.
    /// </summary>
    internal static class StringBuilderExtensions
    {
        /// <summary>
        /// Clears the StringBuilder content.
        /// More readable than: builder.Length = 0;
        /// </summary>
        public static void Clear(this StringBuilder builder)
        {
            builder.Length = 0;
        }

        /// <summary>
        /// Checks if the StringBuilder is empty.
        /// More readable than: builder.Length == 0
        /// </summary>
        public static bool IsEmpty(this StringBuilder builder)
        {
            return builder.Length == 0;
        }

        /// <summary>
        /// Checks if the StringBuilder has content.
        /// More readable than: builder.Length > 0
        /// </summary>
        public static bool HasContent(this StringBuilder builder)
        {
            return builder.Length > 0;
        }

        /// <summary>
        /// Gets the last character without removing it.
        /// Returns '\0' if empty.
        /// </summary>
        public static char LastChar(this StringBuilder builder)
        {
            if (builder.Length == 0)
                return '\0';

            return builder[builder.Length - 1];
        }

        /// <summary>
        /// Checks if the last character is a space.
        /// More readable than: builder[builder.Length - 1] == ' '
        /// </summary>
        public static bool EndsWithSpace(this StringBuilder builder)
        {
            return builder.Length > 0 && builder[builder.Length - 1] == ' ';
        }

        /// <summary>
        /// Removes trailing space if present.
        /// More readable than: if (builder.Length > 0 && builder[builder.Length - 1] == ' ') builder.Length--;
        /// </summary>
        public static void TrimTrailingSpace(this StringBuilder builder)
        {
            if (builder.Length > 0 && builder[builder.Length - 1] == ' ')
            {
                builder.Length--;
            }
        }

        /// <summary>
        /// Appends a space if the builder is not empty.
        /// Useful for word separation.
        /// </summary>
        public static void AppendSpaceIfNotEmpty(this StringBuilder builder)
        {
            if (builder.Length > 0)
            {
                builder.Append(' ');
            }
        }

        /// <summary>
        /// Appends text from a substring without creating intermediate string.
        /// Wraps the 3-parameter Append for better readability.
        /// </summary>
        public static void AppendSubstring(this StringBuilder builder, string text, int startIndex, int length)
        {
            builder.Append(text, startIndex, length);
        }

        /// <summary>
        /// Ensures minimum capacity without excessive allocation.
        /// More readable than: if (builder.Capacity < minCapacity) builder.Capacity = minCapacity;
        /// </summary>
        public static void EnsureCapacity(this StringBuilder builder, int minimumCapacity)
        {
            if (builder.Capacity < minimumCapacity)
            {
                builder.Capacity = minimumCapacity;
            }
        }

        /// <summary>
        /// Returns the content as string and clears the builder.
        /// Useful for "flush" pattern.
        /// </summary>
        public static string ToStringAndClear(this StringBuilder builder)
        {
            string result = builder.ToString();
            builder.Length = 0;
            return result;
        }

        /// <summary>
        /// Adds the current line to a list after trimming trailing space.
        /// Clears the builder afterwards.
        /// Encapsulates the common "finalize line" pattern.
        /// </summary>
        public static void FlushToList(this StringBuilder builder, List<string> lines)
        {
            if (builder.Length > 0)
            {
                builder.TrimTrailingSpace();
                if (builder.Length > 0)
                {
                    lines.Add(builder.ToString());
                }
                builder.Clear();
            }
        }
    }
}