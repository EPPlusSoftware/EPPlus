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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.Rpn.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// This class provides methods for accessing/modifying VBA Functions.
    /// </summary>
    public class FunctionRepository : IFunctionNameProvider
    {
        private Dictionary<Type, FunctionCompiler> _customCompilers = new Dictionary<Type, FunctionCompiler>();
        private Dictionary<Type, RpnFunctionCompiler> _rpnCustomCompilers = new Dictionary<Type, RpnFunctionCompiler>();

        private Dictionary<string, ExcelFunction> _functions = new Dictionary<string, ExcelFunction>(StringComparer.Ordinal);
        
        /// <summary>
        /// Gets a <see cref="Dictionary{Type, FunctionCompiler}" /> of custom <see cref="FunctionCompiler"/>s.
        /// </summary>
        public Dictionary<Type, FunctionCompiler> CustomCompilers
        {
            get { return _customCompilers; }
        }

        internal Dictionary<Type, RpnFunctionCompiler> RpnCustomCompilers
        {
            get { return _rpnCustomCompilers; }
        }

        private FunctionRepository()
        {

        }

        public static FunctionRepository Create()
        {
            var repo = new FunctionRepository();
            repo.LoadModule(new BuiltInFunctions());
            return repo;
        }

        /// <summary>
        /// Loads a module of <see cref="ExcelFunction"/>s to the function repository.
        /// </summary>
        /// <param name="module">A <see cref="IFunctionModule"/> that can be used for adding functions and custom function compilers.</param>
        public virtual void LoadModule(IFunctionModule module)
        {
            foreach (var key in module.Functions.Keys)
            {
                var lowerKey = key.ToLower(CultureInfo.InvariantCulture);
                _functions[lowerKey] = module.Functions[key];
            }
            foreach (var key in module.CustomCompilers.Keys)
            {
                CustomCompilers[key] = module.CustomCompilers[key];
            }
        }

        public virtual ExcelFunction GetFunction(string name)
        {
            if(!_functions.ContainsKey(name.ToLower(CultureInfo.InvariantCulture)))
            {
                //throw new InvalidOperationException("Non supported function: " + name);
                //throw new ExcelErrorValueException("Non supported function: " + name, ExcelErrorValue.Create(eErrorType.Name));
                return null;
            }
            return _functions[name.ToLower(CultureInfo.InvariantCulture)];
        }

        /// <summary>
        /// Removes all functions from the repository
        /// </summary>
        public virtual void Clear()
        {
            _functions.Clear();
        }

        /// <summary>
        /// Returns true if the the supplied <paramref name="name"/> exists in the repository.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public bool IsFunctionName(string name)
        {
            return _functions.ContainsKey(name.ToLower(CultureInfo.InvariantCulture));
        }

        /// <summary>
        /// Returns the names of all implemented functions.
        /// </summary>
        public IEnumerable<string> FunctionNames
        {
            get { return _functions.Keys; }
        }

        /// <summary>
        /// Adds or replaces a function.
        /// </summary>
        /// <param name="functionName"> Case-insensitive name of the function that should be added or replaced.</param>
        /// <param name="functionImpl">An implementation of an <see cref="ExcelFunction"/>.</param>
        public void AddOrReplaceFunction(string functionName, ExcelFunction functionImpl)
        {
            Require.That(functionName).Named("functionName").IsNotNullOrEmpty();
            Require.That(functionImpl).Named("functionImpl").IsNotNull();
            var fName = functionName.ToLower(CultureInfo.InvariantCulture);
            if (_functions.ContainsKey(fName))
            {
                _functions.Remove(fName);
            }
            _functions[fName] = functionImpl;
        }
    }
}
