/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif
namespace OfficeOpenXml.Utils.CompundDocument
{
    internal partial class CompoundDocumentFile 
    {
#if !NET35 && !NET40
        /// <summary>
        /// Verifies that the header is correct.
        /// </summary>
        /// <param name="fi">The file</param>
        /// <returns></returns>
        public static async Task<bool> IsCompoundDocumentAsync(FileInfo fi)
        {
            try
            {
                var fs = fi.OpenRead();
                var b = new byte[8];
                await fs.ReadAsync(b, 0, 8).ConfigureAwait(false);
                return IsCompoundDocument(b);
            }
            catch
            {
                return false;
            }            
        }
        public static async Task<bool> IsCompoundDocumentAsync(MemoryStream ms)
        {
            var pos = ms.Position;
            ms.Position = 0;
            var b=new byte[8];
            await ms.ReadAsync(b, 0, 8).ConfigureAwait(false);
            ms.Position = pos;
            return IsCompoundDocument(b);
        }
#endif
    }
}

