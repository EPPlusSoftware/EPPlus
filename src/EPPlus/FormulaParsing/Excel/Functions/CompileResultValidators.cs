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
    /// Validator collections
    /// </summary>
    internal class CompileResultValidators
    {
        private readonly Dictionary<DataType, CompileResultValidator> _validators = new Dictionary<DataType, CompileResultValidator>(); 

        private CompileResultValidator CreateOrGet(DataType dataType)
        {
            if (_validators.ContainsKey(dataType))
            {
                return _validators[dataType];
            }
            if (dataType == DataType.Decimal)
            {
                return _validators[DataType.Decimal] = new DecimalCompileResultValidator();
            }
            return CompileResultValidator.Empty;
        }
        /// <summary>
        /// Get validator of type
        /// </summary>
        /// <param name="dataType"></param>
        /// <returns></returns>
        public CompileResultValidator GetValidator(DataType dataType)
        {
            return CreateOrGet(dataType);
        }
    }
}
