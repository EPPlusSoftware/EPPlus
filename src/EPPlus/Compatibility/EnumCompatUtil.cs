using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Compatibility
{
    internal static class EnumCompatUtil
    {
        public static bool TryParse<T>(string s, out T result)
            where T : Enum
        {
            result = default(T);
            try
            {
                result = (T)Enum.Parse(typeof(T), s);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
