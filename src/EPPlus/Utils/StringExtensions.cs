namespace OfficeOpenXml.Utils
{
    internal static class StringExtensions
    {
        internal static string NullIfWhiteSpace(this string s) { return s == "" ? null : s; }

        internal static string CapitalizeFirstLetter(this string s) { s = s[0].ToString().ToUpper() + s.Substring(1); return s; }
    }
}
