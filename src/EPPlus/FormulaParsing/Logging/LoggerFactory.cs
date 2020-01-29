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

namespace OfficeOpenXml.FormulaParsing.Logging
{
    /// <summary>
    /// Create loggers that can be used for logging the formula parser.
    /// </summary>
    public static class LoggerFactory
    {
        /// <summary>
        /// Creates a logger that logs to a simple textfile.
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static IFormulaParserLogger CreateTextFileLogger(FileInfo file)
        {
            return new TextFileLogger(file);
        }
    }
}
