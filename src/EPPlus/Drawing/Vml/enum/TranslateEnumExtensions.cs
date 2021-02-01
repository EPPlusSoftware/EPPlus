using System;
namespace OfficeOpenXml.Drawing.Vml
{
    internal static class TranslateEnumExtensions
    {
        internal static eVmlGradientMethod ToGradientMethodEnum(this string s, eVmlGradientMethod defaultValue)
        {
            try
            {
                if (string.IsNullOrEmpty(s)) return defaultValue;
                if(s=="linear sigma")
                {
                    return eVmlGradientMethod.LinearSigma;
                }
                else
                {
                    return (eVmlGradientMethod)Enum.Parse(typeof(eVmlGradientMethod), s, true);
                }                    
            }
            catch
            {
                return defaultValue;
            }
        }
        internal static string ToEnumString(this eVmlGradientMethod enumValue)
        {
            if(enumValue== eVmlGradientMethod.LinearSigma)
            {
                return "linear sigma";
            }
            else
            {
                var s = enumValue.ToString();
                return s.Substring(0, 1).ToLower() + s.Substring(1);
            }
        }
    }
}
