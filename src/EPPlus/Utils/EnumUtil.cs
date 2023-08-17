using System;
using System.Runtime.CompilerServices;

namespace OfficeOpenXml.Utils
{
    internal static class EnumUtil
    {

#if (!NET35)
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
#endif
        public static bool HasFlag<T>(T value, T flag) where T : Enum
        {
            return (Convert.ToInt32(value) & Convert.ToInt32(flag))==Convert.ToInt32(flag);
        }
#if (!NET35)
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
#endif
        public static bool HasNotFlag<T>(T value, T flag) where T : Enum
        {
            return (Convert.ToInt32(value) & Convert.ToInt32(flag)) != Convert.ToInt32(flag);
        }
    }
}
