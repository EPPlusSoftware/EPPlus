using System;

namespace OfficeOpenXml.Utils
{
    internal static class EnumUtil
    {
        public static bool HasFlag<T>(T value, T flag) where T : Enum
        {
            return (Convert.ToInt32(value) & Convert.ToInt32(flag))==Convert.ToInt32(flag);
        }
        public static bool HasNotFlag<T>(T value, T flag) where T : Enum
        {
            return (Convert.ToInt32(value) & Convert.ToInt32(flag)) != Convert.ToInt32(flag);
        }
    }
}
