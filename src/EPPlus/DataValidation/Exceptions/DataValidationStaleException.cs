/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/05/2021         EPPlus Software AB       Added class
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.DataValidation.Exceptions
{
    /// <summary>
    /// This exception is thrown by an instance of <see cref="ExcelDataValidation"/> if it is stale and should not allow setting of properties.
    /// </summary>
    public class DataValidationStaleException : InvalidOperationException
    {

        internal DataValidationStaleException(string message) : base(message)
        {
        }
    }
}
