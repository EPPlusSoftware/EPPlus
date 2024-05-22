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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Information about an argument passed to a function used in the formula parser. 
    /// </summary>
    [Flags]
    public enum FunctionParameterInformation
    { 
        /// <summary>
        /// The argument will be handled as a normally.
        /// </summary>
        Normal = 0x01,
        /// <summary>
        /// If the argument is an address this address will be ignored in the dependency chain.
        /// </summary>
        IgnoreAddress = 0x02,
        /// <summary>
        /// This argument is a condition returning a boolean expression
        /// </summary>
        Condition = 0x04,
        /// <summary>
        /// Use this argument if the condtion is true. Requires a previous parameter to be <see cref="Condition"/>
        /// </summary>
        UseIfConditionIsTrue = 0x08,
        /// <summary>
        /// Use this argument if the condtion is false. Requires a previous parameter to be <see cref="Condition"/>
        /// </summary>
        UseIfConditionIsFalse = 0x10,
        /// <summary>
        /// By default errors found in parameters are returned as a compile result containing the error before calling the <see cref="ExcelFunction.Execute(IList{FunctionArgument}, ParsingContext)"/> method.
        /// Setting this value will allow the function to receive the error as an argument and process them.
        /// </summary>
        IgnoreErrorInPreExecute = 0x20,
        /// <summary>
        /// If the parameter is an address, call the <see cref="ExcelFunction.GetNewParameterAddress"/> to adjust the address before dependency check.
        /// </summary>
        AdjustParameterAddress = 0x40,
        /// <summary>
        /// The parameter is a variable which value is calculated by the next parameter.
        /// </summary>
        IsParameterVariable = 0x80,
    }
}
