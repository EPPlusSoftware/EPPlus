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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Flattens arguments to enumerable
    /// </summary>
    /// <typeparam name="T"></typeparam>
    internal abstract class CollectionFlattener<T>
    {
        /// <summary>
        /// Args to enumerables of type <see cref="IEnumerable{T}"/>
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="convertFunc"></param>
        /// <returns></returns>
        public virtual IEnumerable<T> FuncArgsToFlatEnumerable(IEnumerable<FunctionArgument> arguments, Action<FunctionArgument, IList<T>> convertFunc)
        {
            var argList = new List<T>();
            FuncArgsToFlatEnumerable(arguments, argList, convertFunc);
            return argList;
        }

        private void FuncArgsToFlatEnumerable(IEnumerable<FunctionArgument> arguments, List<T> argList, Action<FunctionArgument, IList<T>> convertFunc)
        {
            foreach (var arg in arguments)
            {
                if (arg.Value is IEnumerable<FunctionArgument>)
                {
                    FuncArgsToFlatEnumerable((IEnumerable<FunctionArgument>)arg.Value, argList, convertFunc);
                }
                else
                {
                    convertFunc(arg, argList);
                }
            }
        }
    }
}
