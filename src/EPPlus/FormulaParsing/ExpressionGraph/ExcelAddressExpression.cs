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
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExcelAddressExpression : AtomicExpression
    {
        /// <summary>
        /// Gets or sets a value that indicates whether or not to resolve directly to an <see cref="ExcelDataProvider.IRangeInfo"/>
        /// </summary>
        public bool ResolveAsRange { get; set; }

        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ParsingContext _parsingContext;
        private readonly RangeAddressFactory _rangeAddressFactory;
        private readonly bool _negate;

        public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext)
            : this(expression, excelDataProvider, parsingContext, new RangeAddressFactory(excelDataProvider), false)
        {

        }
        public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext, bool negate)
            : this(expression, excelDataProvider, parsingContext, new RangeAddressFactory(excelDataProvider), negate)
        {

        }

        public ExcelAddressExpression(string expression, ExcelDataProvider excelDataProvider, ParsingContext parsingContext, RangeAddressFactory rangeAddressFactory, bool negate)
            : base(expression)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            Require.That(parsingContext).Named("parsingContext").IsNotNull();
            Require.That(rangeAddressFactory).Named("rangeAddressFactory").IsNotNull();
            _excelDataProvider = excelDataProvider;
            _parsingContext = parsingContext;
            _rangeAddressFactory = rangeAddressFactory;
            _negate = negate;
        }

        public override bool IsGroupedExpression
        {
            get { return false; }
        }

        public override CompileResult Compile()
        {
            //if (ParentIsLookupFunction)
            //{
            //    return new CompileResult(ExpressionString, DataType.ExcelAddress);
            //}
            //else
            //{
            //    return CompileRangeValues();
            //}
            var cache = _parsingContext.AddressCache;
            var cacheId = cache.GetNewId();
            if(!cache.Add(cacheId, ExpressionString))
            {
                throw new InvalidOperationException("Catastropic error occurred, address caching failed");
            }
            var compileResult = CompileRangeValues();
            compileResult.ExcelAddressReferenceId = cacheId;
            return compileResult;
        }

        private CompileResult CompileRangeValues()
        {
            var c = this._parsingContext.Scopes.Current;
            var resultRange = _excelDataProvider.GetRange(c.Address.Worksheet, c.Address.FromRow, c.Address.FromCol, ExpressionString);
            if (resultRange == null)
            {
                return CompileResult.Empty;
            }
            if (this.ResolveAsRange || resultRange.Address.Rows > 1 || resultRange.Address.Columns > 1)
            {
                return new CompileResult(resultRange, DataType.Enumerable);
            }
            else
            {
                return CompileSingleCell(resultRange);
            }
        }

        private CompileResult CompileSingleCell(ExcelDataProvider.IRangeInfo result)
        {
            var cell = result.FirstOrDefault();
            if (cell == null)
                return CompileResult.Empty;
            var factory = new CompileResultFactory();
            var compileResult = factory.Create(cell.Value);
            if (_negate && compileResult.IsNumeric)
            {
                compileResult = new CompileResult(compileResult.ResultNumeric * -1, compileResult.DataType);
            }
            compileResult.IsHiddenCell = cell.IsHiddenRow;
            return compileResult;
        }
    }
}
