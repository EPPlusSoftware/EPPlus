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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Logging;
using OfficeOpenXml.FormulaParsing.Utilities;
using System.Diagnostics;
using OfficeOpenXml.FormulaParsing.Exceptions;

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
        /// <param name="excelDataProvider">An instance of <see cref="ExcelDataProvider"/> which provides access to a workbook</param>
        public FormulaParser(ExcelDataProvider excelDataProvider)
            : this(excelDataProvider, ParsingContext.Create())
        {
           
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="excelDataProvider">An <see cref="ExcelDataProvider"></see></param>
        /// <param name="parsingContext">Parsing context</param>
        public FormulaParser(ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
        {
            parsingContext.Parser = this;
            parsingContext.ExcelDataProvider = excelDataProvider;
            parsingContext.NameValueProvider = new EpplusNameValueProvider(excelDataProvider);
            parsingContext.RangeAddressFactory = new RangeAddressFactory(excelDataProvider);
            _parsingContext = parsingContext;
            _excelDataProvider = excelDataProvider;
            Configure(configuration =>
            {
                configuration
                    .SetLexer(new Lexer(_parsingContext.Configuration.FunctionRepository, _parsingContext.NameValueProvider))
                    .SetGraphBuilder(new ExpressionGraphBuilder(excelDataProvider, _parsingContext))
                    .SetExpresionCompiler(new ExpressionCompiler())
                    .FunctionRepository.LoadModule(new BuiltInFunctions());
            });
        }

        /// <summary>
        /// This method enables configuration of the formula parser.
        /// </summary>
        /// <param name="configMethod">An instance of the </param>
        public void Configure(Action<ParsingConfiguration> configMethod)
        {
            configMethod.Invoke(_parsingContext.Configuration);
            _lexer = _parsingContext.Configuration.Lexer ?? _lexer;
            _graphBuilder = _parsingContext.Configuration.GraphBuilder ?? _graphBuilder;
            _compiler = _parsingContext.Configuration.ExpressionCompiler ?? _compiler;
        }

        private ILexer _lexer;
        private IExpressionGraphBuilder _graphBuilder;
        private IExpressionCompiler _compiler;

        internal ILexer Lexer { get { return _lexer; } }
        internal IEnumerable<string> FunctionNames { get { return _parsingContext.Configuration.FunctionRepository.FunctionNames; } } 

        internal virtual object Parse(string formula, RangeAddress rangeAddress)
        {
            using (var scope = _parsingContext.Scopes.NewScope(rangeAddress))
            {
                var tokens = _lexer.Tokenize(formula);
                var graph = _graphBuilder.Build(tokens);
                if (graph.Expressions.Count() == 0)
                {
                    return null;
                }
                return _compiler.Compile(graph.Expressions).Result;
            }
        }

        internal virtual object Parse(IEnumerable<Token> tokens, string worksheet, string address)
        {
            var rangeAddress = _parsingContext.RangeAddressFactory.Create(address);
            using (var scope = _parsingContext.Scopes.NewScope(rangeAddress))
            {
                var graph = _graphBuilder.Build(tokens);
                if (graph.Expressions.Count() == 0)
                {
                    return null;
                }
                return _compiler.Compile(graph.Expressions).Result;
            }
        }
        internal virtual object ParseCell(IEnumerable<Token> tokens, string worksheet, int row, int column)
        {
            var rangeAddress = _parsingContext.RangeAddressFactory.Create(worksheet, column, row);
            using (var scope = _parsingContext.Scopes.NewScope(rangeAddress))
            {
                //    _parsingContext.Dependencies.AddFormulaScope(scope);
                var graph = _graphBuilder.Build(tokens);
                if (graph.Expressions.Count() == 0)
                {
                    return 0d;
                }
                try
                {
                    var compileResult = _compiler.Compile(graph.Expressions);
                    // quick solution for the fact that an excelrange can be returned.
                    var rangeInfo = compileResult.Result as ExcelDataProvider.IRangeInfo;
                    if (rangeInfo == null)
                    {
                        return compileResult.Result ?? 0d;
                    }
                    else
                    {
                        if (rangeInfo.IsEmpty)
                        {
                            return 0d;
                        }
                        if (!rangeInfo.IsMulti)
                        {
                            return rangeInfo.First().Value ?? 0d;
                        }
                        // ok to return multicell if it is a workbook scoped name.
                        if (string.IsNullOrEmpty(worksheet))
                        {
                            return rangeInfo;
                        }
                        if (_parsingContext.Debug)
                        {
                            var msg = string.Format("A range with multiple cell was returned at row {0}, column {1}",
                                row, column);
                            _parsingContext.Configuration.Logger.Log(_parsingContext, msg);
                        }
                        return ExcelErrorValue.Create(eErrorType.Value);
                    }
                }
                catch(ExcelErrorValueException ex)
                {
                    if (_parsingContext.Debug)
                    {
                        _parsingContext.Configuration.Logger.Log(_parsingContext, ex);
                    }
                    return ex.ErrorValue;
                }
            }
        }

        /// <summary>
        /// Parses a formula at a specific address
        /// </summary>
        /// <param name="formula">A string containing the formula</param>
        /// <param name="address">Address of the formula</param>
        /// <returns></returns>
        public virtual object Parse(string formula, string address)
        {
            return Parse(formula, _parsingContext.RangeAddressFactory.Create(address));
        }
        
        /// <summary>
        /// Parses a formula
        /// </summary>
        /// <param name="formula">A string containing the formula</param>
        /// <returns>The result of the calculation</returns>
        public virtual object Parse(string formula)
        {
            return Parse(formula, RangeAddress.Empty);
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
            return ParseAt(rangeAddress.Worksheet, rangeAddress.FromRow, rangeAddress.FromCol);
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
                return Parse(f, _parsingContext.RangeAddressFactory.Create(worksheetName,col,row));
            }
            //var dataItem = _excelDataProvider.GetRangeValues(address).FirstOrDefault();
            //if (dataItem == null /*|| (dataItem.Value == null && dataItem.Formula == null)*/) return null;
            //if (!string.IsNullOrEmpty(dataItem.Formula))
            //{
            //    return Parse(dataItem.Formula, _parsingContext.RangeAddressFactory.Create(address));
            //}
            //return Parse(dataItem.Value.ToString(), _parsingContext.RangeAddressFactory.Create(address));
        }


        internal void InitNewCalc()
        {
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
