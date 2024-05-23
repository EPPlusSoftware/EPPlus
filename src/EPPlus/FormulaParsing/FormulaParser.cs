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
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// Entry class for the formula calulation engine of EPPlus.
    /// </summary>
    public class FormulaParser : IDisposable
    {
        private readonly ParsingContext _parsingContext;
        private readonly ExcelDataProvider _excelDataProvider;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="package">The package to calculate</param>
        public FormulaParser(ExcelPackage package)
            : this(new EpplusExcelDataProvider(package), package)
        {

        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="excelDataProvider">An instance of <see cref="ExcelDataProvider"/> which provides access to a workbook</param>
        /// <param name="package">The package to calculate</param>
        internal FormulaParser(ExcelDataProvider excelDataProvider, ExcelPackage package = null)
            : this(excelDataProvider, ParsingContext.Create(package))
        {

        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="excelDataProvider">An <see cref="ExcelDataProvider"></see></param>
        /// <param name="parsingContext">Parsing context</param>
        internal FormulaParser(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
        {
            parsingContext.Parser = this;
            parsingContext.ExcelDataProvider = excelDataProvider;
            parsingContext.NameValueProvider = new EpplusNameValueProvider(excelDataProvider);
            parsingContext.RangeAddressFactory = new RangeAddressFactory(excelDataProvider, parsingContext);
            _parsingContext = parsingContext;
            _excelDataProvider = excelDataProvider;
            Configure(configuration =>
            {
                configuration
                    .FunctionRepository.LoadModule(new BuiltInFunctions());
            });
            Tokenizer = new SourceCodeTokenizer(parsingContext.Configuration.FunctionRepository, parsingContext.NameValueProvider);
        }

        /// <summary>
        /// This method enables configuration of the formula parser.
        /// </summary>
        /// <param name="configMethod">An instance of the </param>
        internal void Configure(Action<ParsingConfiguration> configMethod)
        {
            configMethod.Invoke(_parsingContext.Configuration);
            //_lexer = _parsingContext.Configuration.Lexer ?? _lexer;
            //_graphBuilder = _parsingContext.Configuration.GraphBuilder ?? _graphBuilder;
            //_compiler = _parsingContext.Configuration.ExpressionCompiler ?? _compiler;
        }

        //private ILexer _lexer;
        //private IExpressionGraphBuilder _graphBuilder;
        //private IExpressionCompiler _compiler;
        //internal IExpressionGraphBuilder GraphBuilder => _graphBuilder;
        internal ParsingContext ParsingContext => _parsingContext;
        //internal IExpressionCompiler Compiler => _compiler;

        //internal ILexer Lexer { get { return _lexer; } }
        internal ISourceCodeTokenizer Tokenizer { get; private set; }
        internal IEnumerable<string> FunctionNames { get { return _parsingContext.Configuration.FunctionRepository.FunctionNames; } } 

        /// <summary>
        /// Contains information about filters on a workbook's worksheets.
        /// </summary>
        internal FilterInfo FilterInfo { get; private set; }

        internal virtual object Parse(string formula, FormulaCellAddress cell, ExcelCalculationOption options = default)
        {            
            return RpnFormulaExecution.ExecuteFormula(_parsingContext.Package?.Workbook, formula, cell, options ?? new ExcelCalculationOption());
        }

        /// <summary>
        /// Parse with option to not write result to cell but only return it
        /// </summary>
        /// <param name="formula"></param>
        /// <param name="address"></param>
        /// <param name="writeToCell">True if write result to cell false if not</param>
        /// <returns></returns>
        internal virtual object Parse(string formula, string address, bool writeToCell)
        {
            var calcOption = new ExcelCalculationOption();
            calcOption.AllowCircularReferences = true;
            calcOption.FollowDependencyChain = false;
            
            return RpnFormulaExecution.ExecuteFormula(_parsingContext.Package?.Workbook, formula, _parsingContext.RangeAddressFactory.CreateCell(address), calcOption);
        }

        /// <summary>
        /// Parses a formula at a specific address
        /// </summary>
        /// <param name="formula">A string containing the formula</param>
        /// <param name="address">Address of the formula</param>
        /// <param name="options">Calculation options</param>
        /// <returns></returns>
        public virtual object Parse(string formula, string address, ExcelCalculationOption options = default)
        {
            return Parse(formula, _parsingContext.RangeAddressFactory.CreateCell(address), options);
        }
        
        /// <summary>
        /// Parses a formula
        /// </summary>
        /// <param name="formula">A string containing the formula</param>
        /// <returns>The result of the calculation</returns>
        public virtual object Parse(string formula)
        {
            return Parse(formula, new FormulaCellAddress() { Row = -1});
        }

        /// <summary>
        /// Parses a formula in a specific location
        /// </summary>
        /// <param name="address">address of the cell to calculate</param>
        /// <returns>The result of the calculation</returns>
        public virtual object ParseAt(string address)
        {
            Require.That(address).Named("address").IsNotNullOrEmpty();
            var rangeAddress = _parsingContext.RangeAddressFactory.Create(address);
            return ParseAt(rangeAddress.WorksheetName, rangeAddress.FromRow, rangeAddress.FromCol);
        }

        /// <summary>
        /// Parses a formula in a specific location
        /// </summary>
        /// <param name="worksheetName">Name of the worksheet</param>
        /// <param name="row">Row in the worksheet</param>
        /// <param name="col">Column in the worksheet</param>
        /// <returns>The result of the calculation</returns>
        public virtual object ParseAt(string worksheetName, int row, int col)
        {
            var f = _excelDataProvider.GetRangeFormula(worksheetName, row, col);
            if (string.IsNullOrEmpty(f))
            {
                return _excelDataProvider.GetRangeValue(worksheetName, row, col);
            }
            else
            {
                var wsIx=_parsingContext.GetWorksheetIndex(worksheetName);
                return Parse(f, new FormulaCellAddress(wsIx, row, col));
            }
        }


        internal void InitNewCalc(FilterInfo filterInfo)
        {
            FilterInfo = filterInfo;
            if(_excelDataProvider!=null)
            {
                _excelDataProvider.Reset();
            }
        }

        /// <summary>
        /// An <see cref="IFormulaParserLogger"/> for logging during calculation
        /// </summary>
        public IFormulaParserLogger Logger
        {
            get { return _parsingContext.Configuration.Logger; }
        }

        /// <summary>
        /// Implementation of <see cref="IDisposable"></see>
        /// </summary>
        public void Dispose()
        {
            if (_parsingContext.Debug)
            {
                _parsingContext.Configuration.Logger.Dispose();
            }
        }
	}
}
