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
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
    /// <summary>
    /// Operator interface
    /// </summary>
    internal interface IOperator
    {
        /// <summary>
        /// Operator
        /// </summary>
        Operators Operator { get; }

        /// <summary>
        /// Apply
        /// </summary>
        /// <param name="left"></param>
        /// <param name="right"></param>
        /// <param name="ctx"></param>
        /// <returns></returns>
        CompileResult Apply(CompileResult left, CompileResult right, ParsingContext ctx);

        /// <summary>
        /// Precedence
        /// </summary>
        int Precedence { get; }
    }
}
