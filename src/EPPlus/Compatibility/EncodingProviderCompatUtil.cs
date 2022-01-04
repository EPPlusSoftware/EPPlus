using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Compatibility
{
    /// <summary>
    /// Returns the encoding with the specified code page number
    /// </summary>
    internal static class EncodingProviderCompatUtil
    {
        public static Encoding GetEncoding(int codePage)
        {
#if NETFULL
            return Encoding.GetEncoding(codePage);
#else
            return CodePagesEncodingProvider.Instance.GetEncoding(codePage);
#endif
        }

        /// <summary>
        /// Returns the encoding with the specified name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Encoding GetEncoding(string name)
        {
#if NETFULL
            return Encoding.GetEncoding(name);
#else
            return CodePagesEncodingProvider.Instance.GetEncoding(name);
#endif
        }
    }
}
