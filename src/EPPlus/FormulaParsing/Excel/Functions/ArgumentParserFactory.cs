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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Parser factory for 
    /// </summary>
    internal class ArgumentParserFactory
    {
        /// <summary>
        /// Create argument parser for datatypes <see cref="DataType.Integer"></see>, <see cref="DataType.Boolean"></see> and <see cref="DataType.Decimal"></see>
        /// </summary>
        /// <param name="dataType"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        public virtual ArgumentParser CreateArgumentParser(DataType dataType)
        {
            switch (dataType)
            {
                case DataType.Integer:
                    return new IntArgumentParser();
                case DataType.Boolean:
                    return new BoolArgumentParser();
                case DataType.Decimal:
                    return new DoubleArgumentParser();
                default:
                    throw new InvalidOperationException("non supported argument parser type " + dataType.ToString());
            }
        }
    }
}
