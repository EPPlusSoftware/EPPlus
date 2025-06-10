using OfficeOpenXml.VBA;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Utils.String
{
    internal static class StringExtensions
    {
        internal static string NullIfWhiteSpace(this string s) { return s == "" ? null : s; }

        internal static string CapitalizeFirstLetter(this string s) { s = s[0].ToString().ToUpper() + s.Substring(1); return s; }

        internal static string UnCapitalizeFirstLetter(this string s) { s = s[0].ToString().ToLower() + s.Substring(1); return s; }

        internal static string GetSubstringStoppingAtSymbol(this string s, int index, string stopSymbol = "\"")
        {
            if (!string.IsNullOrEmpty(s))
            {
                int charIndex = s.IndexOf(stopSymbol, index);

                if (charIndex > 0)
                {
                    return s.Substring(index, charIndex - index);
                }
            }

            return string.Empty;
        }
        internal static bool ContainsOnlyCharacter(this string s, char theCharacter, bool ignoreCase = true)
        {
            if (string.IsNullOrEmpty(s)) return false;
            if (ignoreCase)
            {
                s = s.ToLower();
                theCharacter = char.ToLower(theCharacter);
            }

            foreach (var c in s)
            {
                if (c != theCharacter)
                {
                    return false;
                }
            }
            return true;
        }

        internal static string UnitcodeSubstring(this string s, int length)
        {
            return StringUtil.UnicodeSubstring(s, length);
        }

        internal static string UnitcodeSubstring(this string s, int start, int length)
        {
            return StringUtil.UnicodeSubstring(s, start, length);
        }
    }
}
