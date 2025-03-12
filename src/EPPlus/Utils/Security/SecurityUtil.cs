using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OfficeOpenXml.Utils.Security
{
    internal static class SecurityUtil
    {
        /// <summary>
        /// Create a cryptographically strong guid
        /// </summary>
        /// <returns></returns>
        internal static Guid CreateSecureGuid()
        {
            var bytes = new byte[16];
            var rng = RandomNumberGenerator.Create();
            rng.GetBytes(bytes);
            var aGuid = new Guid(bytes);

            return aGuid;
        }
    }
}
