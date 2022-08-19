/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/19/2022         EPPlus Software AB       Implementing handling of initialization errors in ExcelPackage class.
 *************************************************************************************************/
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// This class represents an error/Exception that has occured during initalization.
    /// </summary>
    public class ExcelInitializationError
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="errorMessage"></param>
        /// <param name="e"></param>
        internal ExcelInitializationError(string errorMessage, Exception e)
        {
            Require.Argument(errorMessage).IsNotNullOrEmpty("errorMessage");
            Require.Argument(e).IsNotNull("e");
            ErrorMessage = errorMessage;
            Exception = e;
            TimestampUtc = DateTime.UtcNow;
        }

        private ExcelInitializationError()
        {

        }

        /// <summary>
        /// Error message describing the initialization error
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Timestamp representing when the error occurred
        /// </summary>
        public DateTime TimestampUtc { get; private set; }

        /// <summary>
        /// The <see cref="Exception"/>
        /// </summary>
        public Exception Exception { get; private set; }
    }
}
