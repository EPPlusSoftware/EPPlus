/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  22/11/2024         EPPlus Software AB           EPPlus v8
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Interfaces.Net
{
    /// <summary>
    /// Interface for a service that downloads a byte array via an url.
    /// </summary>
    public interface IHttpsService
    {
        /// <summary>
        /// Returns a <see cref="byte[]"/> via a the <paramref name="url"/>.
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        byte[] Download(string url);
    }
}
