using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType
{
    internal static class StringBuilderExtensions
    {
        /// <summary>
        /// If the last char in the stringbuilder is a space
        /// trim it by adjusting length
        /// </summary>
        /// <param name="sb"></param>
        public static void TrimEndSpace(this StringBuilder sb)
        {
            if (sb.Length > 0 && sb[sb.Length - 1] == ' ')
            {
                sb.Length--;
            }
        }
        /// <summary>
        /// If there is anything in the stringbuilder append space
        /// </summary>
        /// <param name="sb"></param>
        public static void AppendSpace(this StringBuilder sb)
        {
            if (sb.Length > 0)
            {
                sb.Append(' ');
            }
        }
    }
}
