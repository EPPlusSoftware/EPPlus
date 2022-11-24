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
using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.FormulaParsing.Utilities;
namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Provides access to various functionality regarding 
    /// excel formula evaluation.
    /// </summary>
    public class FormulaParserManager
    {
        private readonly FormulaParser _parser;

        internal FormulaParserManager(FormulaParser parser)
        {
            Require.That(parser).Named("parser").IsNotNull();
            _parser = parser;
        }

        /// <summary>
        /// Loads a module containing custom functions to the formula parser. By using
        /// this method you can add your own implementations of Excel functions, by
        /// implementing a <see cref="IFunctionModule"/>.
        /// </summary>
        /// <param name="module">A <see cref="IFunctionModule"/> containing <see cref="ExcelFunction"/>s.</param>
        public void LoadFunctionModule(IFunctionModule module)
        {
            _parser.Configure(x => x.FunctionRepository.LoadModule(module));
        }

        /// <summary>
        /// If the supplied <paramref name="functionName"/> does not exist, the supplied
        /// <paramref name="functionImpl"/> implementation will be added to the formula parser.
        /// If it exists, the existing function will be replaced by the supplied <paramref name="functionImpl">function implementation</paramref>
        /// </summary>
        /// <param name="functionName"></param>
        /// <param name="functionImpl"></param>
        public void AddOrReplaceFunction(string functionName, ExcelFunction functionImpl)
        {
            _parser.Configure(x => x.FunctionRepository.AddOrReplaceFunction(functionName, functionImpl));
        }

        /// <summary>
        /// Copies existing <see cref="ExcelFunction"/>´s from one workbook to another.
        /// </summary>
        /// <param name="otherWorkbook">The workbook containing the forumulas to be copied.</param>
        public void CopyFunctionsFrom(ExcelWorkbook otherWorkbook)
        {
            var functions = otherWorkbook.FormulaParserManager.GetImplementedFunctions();
            foreach (var func in functions)
            {
                AddOrReplaceFunction(func.Key, func.Value);
            }
        }

        /// <summary>
        /// Returns an enumeration of the names of all functions implemented, both the built in functions
        /// and functions added using the LoadFunctionModule method of this class.
        /// </summary>
        /// <returns>Function names in lower case</returns>
        public IEnumerable<string> GetImplementedFunctionNames()
        {
            var fnList = _parser.FunctionNames.ToList();
            fnList.Sort((x, y) => String.Compare(x, y, System.StringComparison.Ordinal));
            return fnList;
        }

        /// <summary>
        /// Returns an enumeration of all implemented functions, including the implementing <see cref="ExcelFunction"/> instance.
        /// </summary>
        /// <returns>An enumeration of <see cref="KeyValuePair{String,ExcelFunction}"/>, where the key is the function name</returns>
        public IEnumerable<KeyValuePair<string, ExcelFunction>> GetImplementedFunctions()
        {
            var functions = new List<KeyValuePair<string, ExcelFunction>>();
            _parser.Configure(parsingConfiguration =>
            {
                foreach (var name in parsingConfiguration.FunctionRepository.FunctionNames)
                {
                    functions.Add(new KeyValuePair<string, ExcelFunction>(name, parsingConfiguration.FunctionRepository.GetFunction(name)));
                }
            });
            return functions;
        }

        /// <summary>
        /// Parses the supplied <paramref name="formula"/> and returns the result.
        /// </summary>
        /// <param name="formula">The formula to parse</param>
        /// <returns>The result of the parsed formula</returns>
        public object Parse(string formula)
        {
            return _parser.Parse(formula);
        }

        /// <summary>
        /// Parses the supplied <paramref name="formula"/> and returns the result.
        /// </summary>
        /// <param name="formula">The formula to parse</param>
        /// <param name="address">The full address in the workbook where the <paramref name="formula"/> should be parsed. Example: you might want to parse the formula of a conditional format, then this should be the address of the cell where the conditional format resides.</param>
        /// <returns>The result of the parsed formula</returns>
        public object Parse(string formula, string address)
        {
            return _parser.Parse(formula, address);
        }

        /// <summary>
        /// Attaches a logger to the <see cref="FormulaParser"/>.
        /// </summary>
        /// <param name="logger">An instance of <see cref="IFormulaParserLogger"/></param>
        /// <see cref="OfficeOpenXml.FormulaParsing.Logging.LoggerFactory"/>
        public void AttachLogger(IFormulaParserLogger logger)
        {
            _parser.Configure(c => c.AttachLogger(logger));
        }

        /// <summary>
        /// Attaches a logger to the formula parser that produces output to the supplied logfile.
        /// </summary>
        /// <param name="logfile"></param>
        public void AttachLogger(FileInfo logfile)
        {
            _parser.Configure(c => c.AttachLogger(LoggerFactory.CreateTextFileLogger(logfile)));
        }
        /// <summary>
        /// Detaches any attached logger from the formula parser.
        /// </summary>
        public void DetachLogger()
        {
            _parser.Configure(c => c.DetachLogger());
        }

        public IEnumerable<IFormulaCellInfo> GetCalculationChain(ExcelRangeBase range)
        {
            Require.That(range).IsNotNull();
            return GetCalculationChain(range, null);
        }

        public IEnumerable<IFormulaCellInfo> GetCalculationChain(ExcelRangeBase range, ExcelCalculationOption options)
        {
            Require.That(range).IsNotNull();
            Init(range.Worksheet.Workbook);
            var filterInfo = new FilterInfo(range.Worksheet.Workbook);
            _parser.InitNewCalc(filterInfo);
            var opt = options != null ? options : new ExcelCalculationOption();
            var dc = DependencyChainFactory.Create(range, opt);
            var result = new List<IFormulaCellInfo>();
            foreach(var co in dc.CalcOrder)
            {
                var fc = dc.list[co];
                var adr = new ExcelAddress(fc.Row, fc.Column, fc.Row, fc.Column);
                var fi = new FormulaCellInfo(fc.ws.Name, adr.Address, fc.Formula);
                result.Add(fi);
            }
            return result;
        }

        private static void Init(ExcelWorkbook workbook)
        {
            workbook._formulaTokens = new CellStore<IList<Token>>(); ;
            foreach (var ws in workbook.Worksheets)
            {
                if (!(ws is ExcelChartsheet))
                {
                    if (ws._formulaTokens != null)
                    {
                        ws._formulaTokens.Dispose();
                    }
                    ws._formulaTokens = new CellStore<IList<Token>>();
                }
            }
        }
    }
}
