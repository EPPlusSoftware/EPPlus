/*******************************************************************************
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
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;

namespace EPPlusTest.FormulaParsing.TestHelpers
{
    public static class FunctionsHelper
    {
        public static IList<FunctionArgument> CreateArgs(params object[] args)
        {
            var list = new List<FunctionArgument>();
            foreach (var arg in args)
            {
                list.Add(new FunctionArgument(arg));
            }
            return list;
        }

        public static IList<FunctionArgument> Empty()
        {
            return new List<FunctionArgument>() {new FunctionArgument(null, DataType.Empty)};
        }
    }
}
