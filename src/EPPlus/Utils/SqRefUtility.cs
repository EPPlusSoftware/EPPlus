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
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Utils
{
    /// <summary>
    /// Class for handling translation between ExcelAddresses and sqref addresses.
    /// </summary>
    public static class SqRefUtility
    {
       /// <summary>
       /// Transforms an address to a valid sqRef address.
       /// </summary>
       /// <param name="address">The address to transform</param>
       /// <returns>A valid SqRef address</returns>
       public static string ToSqRefAddress(string address)
       {
           Require.Argument(address).IsNotNullOrEmpty(address);
           address = address.Replace(",", " ");
           address = new Regex("[ ]+").Replace(address, " ");
           return address;
       }

       /// <summary>
       /// Transforms an sqRef address into a excel address
       /// </summary>
       /// <param name="address">The address to transform</param>
       /// <returns>A valid excel address</returns>
       public static string FromSqRefAddress(string address)
       {
           Require.Argument(address).IsNotNullOrEmpty(address);
           address = address.Replace(" ", ",");
           return address;
       }
    }
}
