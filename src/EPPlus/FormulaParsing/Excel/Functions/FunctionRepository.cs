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
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.FormulaExpressions.FunctionCompilers;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// This class provides methods for accessing/modifying VBA Functions.
    /// </summary>
    public class FunctionRepository : IFunctionNameProvider
    {
        private Dictionary<string, ExcelFunction> _functions = new Dictionary<string, ExcelFunction>(StringComparer.Ordinal);
        
        private FunctionRepository()
        {

        }
        /// <summary>
        /// Create repository
        /// </summary>
        /// <returns></returns>
        public static FunctionRepository Create()
        {
            var repo = new FunctionRepository();
            repo.LoadModule(new BuiltInFunctions());
            repo._namespaceFunctions = null;
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
            //foreach (var key in module.CustomCompilers.Keys)
            //{
            //    CustomCompilers[key] = module.CustomCompilers[key];
            //}
            _namespaceFunctions = null;
        }
        /// <summary>
        /// Get function
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
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
            _namespaceFunctions = null;
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
        internal Dictionary<string, string> _namespaceFunctions = null;
        /// <summary>
        /// Contains all functions that needs a namespace prefix in Excel.
        /// For example: The Filter function must have the prefix "_xlfn._xlws."
        /// </summary>
        public Dictionary<string, string> NamespaceFunctions
        {
            get
            {
                if (_namespaceFunctions == null)
                {
                    _namespaceFunctions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var f in _functions)
                    {
                        if (string.IsNullOrEmpty(f.Value.NamespacePrefix) == false)
                        {
                            _namespaceFunctions.Add(f.Key, f.Value.NamespacePrefix);
                        }
                    }
                }
                return _namespaceFunctions;
            }
        }
    }
}
